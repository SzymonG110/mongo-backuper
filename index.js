const mongoose = require('mongoose');
const inquirer = require('inquirer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

(async () => {
    const prompt = inquirer.createPromptModule();

    const action = await prompt([{
        type: 'list',
        name: 'action',
        message: 'Choose an action:',
        choices: ['Create Backup', 'Load Backup']
    }]);

    const mongoURI = (await prompt([{
        type: 'input',
        name: 'mongoURI',
        message: 'Enter MongoDB URI:',
        default: 'mongodb://localhost:27017/'
    }])).mongoURI;

    await mongoose.connect(mongoURI);

    if (action.action === 'Create Backup') {
        await createBackup();
    } else {
        await loadBackup();
    }

    await mongoose.disconnect();
})();

const createBackup = async () => {
    const dbs = (await mongoose.connection.db.admin().listDatabases())
        .databases
        .map(db => db.name)
        .filter(db => !['admin', 'config', 'local'].includes(db));

    const selectedDb = (await inquirer.createPromptModule()([{
        type: 'list',
        name: 'db',
        message: 'Select a database to back up:',
        choices: ["*", ...dbs]
    }])).db;

    if (selectedDb !== '*') {
        const db = mongoose.connection.useDb(selectedDb);
        await createDatabaseBackup(db);
        console.log(`Data exported to ${selectedDb}.xlsx`);
    } else {
        for (const dbName of dbs) {
            console.log(`>>> Exporting data from ${dbName}...`);
            const db = mongoose.connection.useDb(dbName);
            await createDatabaseBackup(db);
            console.log(`>> Data exported to ${dbName}.xlsx`);
        }
    }
};

const loadBackup = async () => {
    const backupsDir = path.join(__dirname, 'backups');
    const backupFiles = fs.readdirSync(backupsDir).filter(file => file.endsWith('.xlsx'));

    if (backupFiles.length === 0) {
        console.log(">>> No backup files found.");
        return;
    }

    const selectedBackup = (await inquirer.createPromptModule()([{
        type: 'list',
        name: 'backup',
        message: 'Select a backup file to load:',
        choices: backupFiles
    }])).backup;

    const backupPath = path.join(backupsDir, selectedBackup);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(backupPath);

    const dbName = selectedBackup.split('-')[0];
    const db = mongoose.connection.useDb(dbName);

    const collections = workbook.worksheets;

    for (const sheet of collections) {
        const collectionName = sheet.name;
        const rows = sheet.getSheetValues();
        const headers = rows[1];

        const existingCollection = (await db.listCollections({ name: collectionName })).map(collection => collection.name).find(name => name === collectionName);

        if (existingCollection) {
            const { overwrite } = await inquirer.createPromptModule()([{
                type: 'confirm',
                name: 'overwrite',
                message: `Collection ${collectionName} already exists. Overwrite?`,
                default: false
            }]);

            if (!overwrite) {
                continue;
            }
        }

        const data = rows.slice(2).map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                if (header) {
                    const value = row[index] || null;
                    if (header === '_id' && value) {
                        obj[header] = new mongoose.Types.ObjectId(value.replace(/^"|"$/g, ''));
                    } else {
                        obj[header] = value;
                    }
                }
            });
            return obj;
        }).filter(obj => Object.keys(obj).length > 0);

        if (data.length > 0) {
            await db.collection(collectionName).deleteMany({});
            await db.collection(collectionName).insertMany(data);
            console.log(`> Loaded ${data.length} records into ${collectionName}.`);
        }
    }

    console.log(`>> Backup from ${selectedBackup} has been loaded.`);
};

const createDatabaseBackup = async (db) => {
    const collections = (await db.listCollections()).map(collection => collection.name);
    const workbook = new ExcelJS.Workbook();

    for (const collection of collections) {
        const sheet = workbook.addWorksheet(collection);
        const dataCursor = db.collection(collection).find();

        let i = 0;
        for await (const data of dataCursor) {
            if (i++ === 0) {
                sheet.addRow(Object.keys(data)).commit();
            }
            sheet.addRow(Object.values(data)).commit();
        }
    }

    const backupsDir = path.join(__dirname, 'backups');
    if (!fs.existsSync(backupsDir)) {
        fs.mkdirSync(backupsDir);
    }

    const timestamp = new Date().toLocaleString("pl").replaceAll(".", "-").replaceAll(":", "-").replaceAll(",", "").replaceAll(" ", "-");
    await workbook.xlsx.writeFile(path.join(backupsDir, `${db.name}-${timestamp}.xlsx`));
};
