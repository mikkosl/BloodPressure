#include "Database.h"
extern "C" {
#include "sqlite3.h"
}

#include <windows.h>
#include <ctime>
#include <iomanip>
#include <sstream>

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
    // Open with UTF-16 Windows path support
    if (sqlite3_open16(dbPath_.c_str(), &db_) != SQLITE_OK)
    {
        return false;
    }
    return EnsureSchema();
}

bool Database::EnsureSchema()
{
    const char* sql =
        "PRAGMA journal_mode=WAL;"
        "CREATE TABLE IF NOT EXISTS readings ("
        "  id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  ts_utc TEXT NOT NULL,"
        "  systolic INTEGER NOT NULL,"
        "  diastolic INTEGER NOT NULL,"
        "  pulse INTEGER NOT NULL,"
        "  note TEXT"
        ");"
        "CREATE INDEX IF NOT EXISTS idx_readings_ts ON readings(ts_utc);";

    char* errMsg = nullptr;
    int rc = sqlite3_exec(db_, sql, nullptr, nullptr, &errMsg);
    if (rc != SQLITE_OK)
    {
        if (errMsg) sqlite3_free(errMsg);
        return false;
    }
    return true;
}

bool Database::AddReading(int systolic, int diastolic, int pulse, const wchar_t* note)
{
    if (!db_) return false;

    const char* sql =
        "INSERT INTO readings(ts_utc, systolic, diastolic, pulse, note) "
        "VALUES (?, ?, ?, ?, ?);";

    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr) != SQLITE_OK)
        return false;

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

    int rc = sqlite3_step(stmt);
    sqlite3_finalize(stmt);
    return rc == SQLITE_DONE;
}

bool Database::GetReadingCount(int& outCount) const
{
    outCount = 0;
    if (!db_) return false;

    const char* sql = "SELECT COUNT(*) FROM readings;";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db_, sql, -1, &stmt, nullptr) != SQLITE_OK)
        return false;

    int rc = sqlite3_step(stmt);
    if (rc == SQLITE_ROW)
    {
        outCount = sqlite3_column_int(stmt, 0);
        sqlite3_finalize(stmt);
        return true;
    }

    sqlite3_finalize(stmt);
    return false;
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