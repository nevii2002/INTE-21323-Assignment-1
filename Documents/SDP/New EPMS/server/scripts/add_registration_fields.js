const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.join(__dirname, '../database.sqlite');
const db = new sqlite3.Database(dbPath);

console.log("Opening DB at:", dbPath);

db.serialize(() => {
    db.run("ALTER TABLE Users ADD COLUMN mobileNumber VARCHAR(255)", (err) => {
        if (err) console.log("mobileNumber error:", err.message);
        else console.log("mobileNumber added.");
    });
    db.run("ALTER TABLE Users ADD COLUMN position VARCHAR(255)", (err) => {
        if (err) console.log("position error:", err.message);
        else console.log("position added.");
    });
});
