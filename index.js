const express = require("express");
const { google } = require("googleapis");
const app = express();

app.get("/", async (req, res) => {

    res.send("<h1>Welcome to the server!</h1><p>Access <a href='https://docs.google.com/spreadsheets/d/1Wl7zBaFEboIx7EMJin1gYtIcNdBCo7Kb1y9dlW9vtgM/edit?usp=sharing'>this link</a> to see the spreadsheet.</p><br><p>Keep this page open. Try entering new values in the cells and refresh this page to see the results.</p>");

    try {
        // Creates an authentication object to use the Google Sheets API.
        const auth = new google.auth.GoogleAuth({
            keyFile: "credentials.json",
            scopes: "https://www.googleapis.com/auth/spreadsheets",
        });

        // Auth Client Instance.
        const client = await auth.getClient();

        // Google Sheets API Instance.
        const googleSheets = google.sheets({ version: "v4", auth: client });

        // Spreadsheet ID.
        const spreadsheetId = "1Wl7zBaFEboIx7EMJin1gYtIcNdBCo7Kb1y9dlW9vtgM";

        // Get Grades from D4 to F column.
        const getStudentsGrades = await googleSheets.spreadsheets.values.get({
            auth,
            spreadsheetId,
            range: "engenharia_de_software!D4:F",
        });

        // Get number of Absences from C column.
        const getStudentsAbsences = await googleSheets.spreadsheets.values.get({
            auth,
            spreadsheetId,
            range: "engenharia_de_software!C4:C",
        });

        // Extracts student grades and absences data from the response obtained.
        const values = getStudentsGrades.data.values;
        const absencesValues = getStudentsAbsences.data.values;
        console.log('Student data obtained successfully.');

        if (values.length && absencesValues) {
            // State Calc for each student.
            const states = values.map((row, index) => {
                const exam1 = parseFloat(row[0]);
                const exam2 = parseFloat(row[1]);
                const exam3 = parseFloat(row[2]);
                const mean = (exam1 + exam2 + exam3) / 3;
                const absences = parseFloat(absencesValues[index][0]);

                let state;
                if (absences > 15) {
                    state = "Reprovado por Falta";
                } else if (mean < 50) {
                    state = "Reprovado por Nota";
                } else if (mean >= 50 && mean < 70) {
                    state = "Exame Final";
                } else {
                    state = "Aprovado";
                }

                return [state, mean];
            });

            // Update State Data on G Column (Starting from G4).
            const writeState = await googleSheets.spreadsheets.values.update({
                auth,
                spreadsheetId,
                range: "engenharia_de_software!G4",
                valueInputOption: 'RAW',
                resource: { values: states }
            });

            console.log('Grades calculated and current situation updated.');
            await calculateFinalExam(auth, googleSheets, spreadsheetId, states);
        }
    } catch (error) {
        console.error('Error starting server:', error);
        res.send('Error starting server. Please try again later.');
    }
});

// Calc Min. Grade. only for students that have State equal to "Exame Final".
async function calculateFinalExam(auth, googleSheets, spreadsheetId, states) {
    const finalExamStates = states.map(([state, mean]) => {
        if (state === "Exame Final") {
            const examScore = 140 - mean;
            return [Math.ceil(examScore)];
        } else {
            return [0];
        }
    });

    console.log('Final Exam score calculated.');

    // Update Final Exam Data on H Column.
    const writeFinalExam = await googleSheets.spreadsheets.values.update({
        auth,
        spreadsheetId,
        range: "engenharia_de_software!H4",
        valueInputOption: 'RAW',
        resource: { values: finalExamStates }
    });

    console.log('Spreadsheet updated successfully!');

}

app.listen(1337, (req, res) => console.log("Application running on port 1337! Access http://localhost:1337 to proceed!")).on('error', (error) => {
    console.error('Error starting server:', error);
    console.log('Please make sure no other application is running on port 1337.');
});