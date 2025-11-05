#include "Database.h"

extern "C" {
#include "sqlite3.h"
}

#include <windows.h>
#include <ctime>
#include <iomanip>
#include <sstream>

static void LogSqliteError(sqlite3* db, const char* where, int rc = 0)
{
    if (!db) return;
    const char* msg = sqlite3_errmsg(db);
    wchar_t wWhere[128]{};
    wchar_t wMsg[512]{};
    MultiByteToWideChar(CP_UTF8, 0, where, -1, wWhere, 128);
    MultiByteToWideChar(CP_UTF8, 0, msg ? msg : "unknown", -1, wMsg, 512);
    std::wstring out = L"[SQLite] ";
    out += wWhere;
    out += L" (rc=" + std::to_wstring(rc) + L"): ";
    out += wMsg;
    out += L"\r\n";
    OutputDebugStringW(out.c_str());
}

static int ExecWithRetry(sqlite3* db, const char* sql, const char* where, DWORD totalMs = 8000)
{
    DWORD waited = 0;
    DWORD delay = 50;
    char* errMsg = nullptr;

    for (;;)
    {
        int rc = sqlite3_exec(db, sql, nullptr, nullptr, &errMsg);
        if (rc == SQLITE_OK) return SQLITE_OK;

        if (rc == SQLITE_BUSY || rc == SQLITE_LOCKED)
        {
            if (waited >= totalMs) {
                if (errMsg) sqlite3_free(errMsg);
                LogSqliteError(db, where, rc);
                return rc;
            }
            Sleep(delay);
            waited += delay;
            if (delay < 800) delay *= 2;
            if (errMsg) { sqlite3_free(errMsg); errMsg = nullptr; }
            continue;
        }

        if (errMsg) sqlite3_free(errMsg);
        LogSqliteError(db, where, rc);
        return rc;
    }
}

static std::wstring Utf8ToWString(const char* s)
{
    if (!s) return {};
    int needed = MultiByteToWideChar(CP_UTF8, 0, s, -1, nullptr, 0);
    if (needed <= 0) return {};
    std::wstring out(needed - 1, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s, -1, out.data(), needed);
    return out;
}

Database::Database(const wchar_t* dbPath)
    : dbPath_(dbPath ? dbPath : L"BloodPressure.db")
{
}

Database::~Database()
{
    if (db_)
    {
        sqlite3_close(db_);
        db_ = nullptr;
    }
}

bool Database::Initialize()
{
    // Convert path to UTF-8 and open with FULLMUTEX for robustness
    std::string pathUtf8 = WStringToUtf8(dbPath_);
    int rc = sqlite3_open_v2(
        pathUtf8.c_str(),
        &db_,
        SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_FULLMUTEX,
        nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "open_v2", rc);
        if (db_) { sqlite3_close(db_); db_ = nullptr; }
        return false;
    }

    sqlite3_extended_result_codes(db_, 1);
    sqlite3_busy_timeout(db_, 5000);

    // Use WAL so readers (DB Browser) don't block writers
    ExecWithRetry(db_, "PRAGMA journal_mode=WAL;", "PRAGMA journal_mode");
    ExecWithRetry(db_, "PRAGMA synchronous=NORMAL;", "PRAGMA synchronous");
    ExecWithRetry(db_, "PRAGMA wal_autocheckpoint=1000;", "PRAGMA wal_autocheckpoint");
    ExecWithRetry(db_, "PRAGMA temp_store=MEMORY;", "PRAGMA temp_store");

    return EnsureSchema();
}

bool Database::EnsureSchema()
{
    const char* sql =
        "CREATE TABLE IF NOT EXISTS readings ("
        "  id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  ts_utc TEXT NOT NULL,"
        "  systolic INTEGER NOT NULL,"
        "  diastolic INTEGER NOT NULL,"
        "  pulse INTEGER NOT NULL,"
        "  note TEXT"
        ");"
        "CREATE INDEX IF NOT EXISTS idx_readings_ts ON readings(ts_utc);";

    int rc = ExecWithRetry(db_, sql, "EnsureSchema");
    return rc == SQLITE_OK;
}

bool Database::AddReading(int systolic, int diastolic, int pulse, const wchar_t* note)
{
    if (!db_) return false;

    const char* sql =
        "INSERT INTO readings(ts_utc, systolic, diastolic, pulse, note) "
        "VALUES (?, ?, ?, ?, ?);";

    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "prepare(insert)", rc);
        return false;
    }

    const std::string ts = UtcNowIso8601();
    sqlite3_bind_text(stmt, 1, ts.c_str(), (int)ts.size(), SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, systolic);
    sqlite3_bind_int(stmt, 3, diastolic);
    sqlite3_bind_int(stmt, 4, pulse);

    std::string noteUtf8;
    if (note && *note)
    {
        noteUtf8 = WStringToUtf8(std::wstring(note));
        sqlite3_bind_text(stmt, 5, noteUtf8.c_str(), (int)noteUtf8.size(), SQLITE_TRANSIENT);
    }
    else
    {
        sqlite3_bind_null(stmt, 5);
    }

    for (;;)
    {
        rc = sqlite3_step(stmt);
        if (rc == SQLITE_DONE) break;
        if (rc == SQLITE_BUSY || rc == SQLITE_LOCKED) { Sleep(50); continue; }
        LogSqliteError(db_, "step(insert)", rc);
        sqlite3_finalize(stmt);
        return false;
    }

    sqlite3_finalize(stmt);
    return true;
}

bool Database::GetReadingCount(int& outCount) const
{
    outCount = 0;
    if (!db_) return false;

    const char* sql = "SELECT COUNT(*) FROM readings;";
    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "prepare(count)", rc);
        return false;
    }

    rc = sqlite3_step(stmt);
    if (rc == SQLITE_ROW)
    {
        outCount = sqlite3_column_int(stmt, 0);
        sqlite3_finalize(stmt);
        return true;
    }

    LogSqliteError(db_, "step(count)", rc);
    sqlite3_finalize(stmt);
    return false;
}

bool Database::GetRecentReadings(int limit, std::vector<Reading>& out) const
{
    out.clear();
    if (!db_ || limit <= 0) return false;

    const char* sql =
        "SELECT id, ts_utc, systolic, diastolic, pulse, COALESCE(note,'') "
        "FROM readings "
        "ORDER BY ts_utc DESC "
        "LIMIT ?;";
    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "prepare(get recent)", rc);
        return false;
    }

    sqlite3_bind_int(stmt, 1, limit);

    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW)
    {
        Reading r{};
        r.id       = sqlite3_column_int(stmt, 0);
        r.tsUtc    = Utf8ToWString(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1)));
        r.systolic = sqlite3_column_int(stmt, 2);
        r.diastolic= sqlite3_column_int(stmt, 3);
        r.pulse    = sqlite3_column_int(stmt, 4);
        r.note     = Utf8ToWString(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 5)));
        out.push_back(std::move(r));
    }

    if (rc != SQLITE_DONE)
    {
        LogSqliteError(db_, "step(get recent)", rc);
        sqlite3_finalize(stmt);
        return false;
    }

    sqlite3_finalize(stmt);
    return true;
}

bool Database::GetRecentReadingsPage(int limit, int offset, std::vector<Reading>& out) const
{
    out.clear();
    if (!db_ || limit <= 0 || offset < 0) return false;

    const char* sql =
        "SELECT id, ts_utc, systolic, diastolic, pulse, COALESCE(note,'') "
        "FROM readings "
        "ORDER BY ts_utc DESC "
        "LIMIT ? OFFSET ?;";
    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "prepare(get recent page)", rc);
        return false;
    }

    sqlite3_bind_int(stmt, 1, limit);
    sqlite3_bind_int(stmt, 2, offset);

    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW)
    {
        Reading r{};
        r.id        = sqlite3_column_int(stmt, 0);
        r.tsUtc     = Utf8ToWString(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1)));
        r.systolic  = sqlite3_column_int(stmt, 2);
        r.diastolic = sqlite3_column_int(stmt, 3);
        r.pulse     = sqlite3_column_int(stmt, 4);
        r.note      = Utf8ToWString(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 5)));
        out.push_back(std::move(r));
    }

    if (rc != SQLITE_DONE)
    {
        LogSqliteError(db_, "step(get recent page)", rc);
        sqlite3_finalize(stmt);
        return false;
    }

    sqlite3_finalize(stmt);
    return true;
}

bool Database::UpdateReading(int id, int systolic, int diastolic, int pulse, const wchar_t* note)
{
    if (!db_ || id <= 0) return false;

    const char* sql =
        "UPDATE readings SET systolic=?, diastolic=?, pulse=?, note=? WHERE id=?;";
    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK)
    {
        LogSqliteError(db_, "prepare(update)", rc);
        return false;
    }

    sqlite3_bind_int(stmt, 1, systolic);
    sqlite3_bind_int(stmt, 2, diastolic);
    sqlite3_bind_int(stmt, 3, pulse);

    std::string noteUtf8;
    if (note && *note)
    {
        noteUtf8 = WStringToUtf8(std::wstring(note));
        sqlite3_bind_text(stmt, 4, noteUtf8.c_str(), (int)noteUtf8.size(), SQLITE_TRANSIENT);
    }
    else
    {
        sqlite3_bind_null(stmt, 4);
    }

    sqlite3_bind_int(stmt, 5, id);

    for (;;)
    {
        rc = sqlite3_step(stmt);
        if (rc == SQLITE_DONE) break;
        if (rc == SQLITE_BUSY || rc == SQLITE_LOCKED) { Sleep(50); continue; }
        LogSqliteError(db_, "step(update)", rc);
        sqlite3_finalize(stmt);
        return false;
    }

    sqlite3_finalize(stmt);
    return true;
}

bool Database::DeleteReading(int id)
{
    if (!db_) return false;

    const char* sql = "DELETE FROM readings WHERE id = ?";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr) != SQLITE_OK)
        return false;

    if (sqlite3_bind_int(stmt, 1, id) != SQLITE_OK) {
        sqlite3_finalize(stmt);
        return false;
    }

    bool result = (sqlite3_step(stmt) == SQLITE_DONE);
    sqlite3_finalize(stmt);
    return result;
}

std::string Database::WStringToUtf8(const std::wstring& w)
{
    if (w.empty()) return {};
    int needed = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string out(needed, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), out.data(), needed, nullptr, nullptr);
    return out;
}

std::string Database::UtcNowIso8601()
{
    std::time_t now = std::time(nullptr);
    std::tm tmUtc{};
    gmtime_s(&tmUtc, &now);

    std::ostringstream oss;
    oss << std::put_time(&tmUtc, "%Y-%m-%dT%H:%M:%SZ");
    return oss.str();
}