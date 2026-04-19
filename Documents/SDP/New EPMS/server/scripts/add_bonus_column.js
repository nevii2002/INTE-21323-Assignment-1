const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./database.sqlite');

db.serialize(() => {
    db.run("ALTER TABLE Users ADD COLUMN bonusAmount FLOAT DEFAULT 0", (err) => {
        if (err) {
            console.log("Column bonusAmount might already exist or error:", err.message);
        } else {
            console.log("Column bonusAmount added successfully.");
        }
        db.close();
    });
});
