const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./database.sqlite');

db.serialize(() => {
    db.all("SELECT name FROM sqlite_master WHERE type='table'", (err, tables) => {
        if (err) {
            console.error(err);
            db.close();
            return;
        }
        console.log("Tables:", tables.map(t => t.name));

        const backupTable = tables.find(t => t.name === 'Users_backup');
        if (backupTable) {
            console.log("Found Users_backup. Dropping it...");
            db.run("DROP TABLE Users_backup", (err) => {
                if (err) console.error("Error dropping backup:", err);
                else console.log("Dropped Users_backup successfully.");
                db.close();
            });
        } else {
            console.log("Users_backup not found.");
            db.close();
        }
    });
});
