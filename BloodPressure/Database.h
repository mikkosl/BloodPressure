#pragma once
#include <string>
#include <vector>

// Forward declare sqlite3 to avoid leaking its headers
struct sqlite3;

struct Reading
{
    std::wstring tsUtc;
    int systolic{};
    int diastolic{};
    int pulse{};
    std::wstring note;
};

class Database
{
public:
    explicit Database(const wchar_t* dbPath);
    ~Database();

    bool Initialize();
    bool AddReading(int systolic, int diastolic, int pulse, const wchar_t* note);
    bool GetReadingCount(int& outCount) const;

    // New: fetch most recent readings (ordered by timestamp desc)
    bool GetRecentReadings(int limit, std::vector<Reading>& out) const;

private:
    bool EnsureSchema();
    static std::string WStringToUtf8(const std::wstring& w);
    static std::string UtcNowIso8601();

private:
    std::wstring dbPath_;
    sqlite3* db_{ nullptr };
};