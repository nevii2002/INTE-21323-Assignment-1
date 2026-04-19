const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./database.sqlite');

db.serialize(() => {
    db.run("ALTER TABLE EmployeeKPIs ADD COLUMN customBonus FLOAT DEFAULT 0", (err) => {
        if (err) {
            console.log("Column customBonus might already exist or error:", err.message);
        } else {
            console.log("Column customBonus added successfully.");
        }
        db.close();
    });
});
