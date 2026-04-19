const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./database.sqlite');

db.serialize(() => {
    // Attempt to add column. Will fail if exists, which is fine.
    db.run("ALTER TABLE Users ADD COLUMN basicSalary FLOAT DEFAULT 0", (err) => {
        if (err) {
            console.log("Column might already exist or error:", err.message);
        } else {
            console.log("Column basicSalary added successfully.");
        }
        db.close();
    });
});
