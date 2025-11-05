#pragma once
#include <string>
#include <vector>

// Forward declare sqlite3 to avoid leaking its headers
struct sqlite3;

struct Reading
{
    int id{};                 // new: primary key for editing
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
    bool UpdateReading(int id, int systolic, int diastolic, int pulse, const wchar_t* note);
    bool DeleteReading(int id);
    bool DeleteAllReadings();
    bool GetReadingCount(int& outCount) const;

    // New: fetch all readings (ordered by timestamp desc)
    bool GetAllReadings(std::vector<Reading>& out) const; // <-- add

    // New: fetch most recent readings (ordered by timestamp desc)
    bool GetRecentReadings(int limit, std::vector<Reading>& out) const;

    // New: fetch most recent readings with paging (ordered by timestamp desc)
    bool GetRecentReadingsPage(int limit, int offset, std::vector<Reading>& out) const;

private:
    bool EnsureSchema();
    static std::string WStringToUtf8(const std::wstring& w);
    static std::string UtcNowIso8601();

private:
    std::wstring dbPath_;
    sqlite3* db_{ nullptr };    
};

