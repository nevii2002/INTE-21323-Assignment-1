const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const dbPath = path.resolve(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error('Error opening database', err.message);
        return;
    }
    console.log('Connected to the SQLite database.');

    db.run('ALTER TABLE Users ADD COLUMN jobDescription TEXT', (err) => {
        if (err) {
            console.log('Error adding jobDescription:', err.message);
        } else {
            console.log('jobDescription added.');
        }

        db.run('ALTER TABLE Users ADD COLUMN responsibilities TEXT', (err2) => {
            if (err2) {
                console.log('Error adding responsibilities:', err2.message);
            } else {
                console.log('responsibilities added.');
            }
            db.close();
        });
    });
});
