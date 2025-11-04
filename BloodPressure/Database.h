#pragma once
#include <string>

// Forward declare sqlite3 to avoid leaking its headers
struct sqlite3;

class Database
{
public:
    explicit Database(const wchar_t* dbPath);
    ~Database();

    bool Initialize();
    bool AddReading(int systolic, int diastolic, int pulse, const wchar_t* note);
    bool GetReadingCount(int& outCount) const;

private:
    bool EnsureSchema();
    static std::string WStringToUtf8(const std::wstring& w);
    static std::string UtcNowIso8601();

private:
    std::wstring dbPath_;
    sqlite3* db_{ nullptr };
};