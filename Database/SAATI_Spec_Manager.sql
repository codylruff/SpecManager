-- warping specifications table
CREATE TABLE IF NOT EXISTS warping_specifications (
    Id                          INTEGER PRIMARY KEY,
    Material_Id                 TEXT NOT NULL DEFAULT "",
    Time_Stamp                  DATETIME NOT NULL,
    Json_Text                   TEXT NOT NULL DEFAULT "{}",
    SpecType                TEXT NOT NULL DEFAULT ""
)

-- style specifications table
CREATE TABLE IF NOT EXISTS style_specifications (
    Id                      INTEGER PRIMARY KEY,
    Material_Id             TEXT NOT NULL DEFAULT "",
    Time_Stamp              DATETIME NOT NULL,
    Json_Text               TEXT NOT NULL DEFAULT "{}",
    SpecType                TEXT NOT NULL DEFAULT ""
)

-- template_specifications table
CREATE TABLE IF NOT EXISTS template_specifications (
    Id                      INTEGER PRIMARY KEY,
    SpecType                TEXT NOT NULL DEFAULT "",
    Time_Stamp              TEXT NOT NULL,
    Json_Text               TEXT NOT NULL DEFAULT "",
    Material_Id             TEXT NOT NULL DEFAULT ""
)
