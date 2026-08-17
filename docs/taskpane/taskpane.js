/* global Office, OfficeRuntime */

let initialized = false;
let currentUserEmail = "";
const SHEET_PASSWORD = "isa2026";

Office.onReady(() => {
    if (initialized) return;
    initialized = true;

    document.getElementById("status").textContent = "Office is ready";
    document.getElementById("userEmail").textContent = "Identifying you...";

    document.getElementById("sendtoalt").addEventListener(
        "click",
        () => sendToAlternate(currentUserEmail)
    );

    document.getElementById("alternateEmail").addEventListener(
        "input",
        function () {

            const hasAlternate =
                this.value.trim() !== "";

            document.getElementById("sendtoalt").disabled =
            !hasAlternate;

            document.getElementById("submitVote").disabled =
            hasAlternate;
        }
    );

    document.getElementById("sendtoalt").disabled = true;
    document.getElementById("submitVote").disabled = false;

    OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true
    })
    .then(token => {
        const payload = JSON.parse(
            atob(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))
        );

        const email =
            payload.preferred_username ||
            payload.upn ||
            payload.email;
            currentUserEmail = email;

        document.getElementById("userEmail").textContent =
            email || "Email not found";

        document.getElementById("status").textContent =
            "Authentication successful";

        console.log("User:", email);

        // Wait for authentication to finish before touching Excel.
        setTimeout(() => findUser(email), 1000);
    })
    .catch(error => {
        document.getElementById("userEmail").textContent =
            "Authentication failed";

        document.getElementById("status").textContent =
            error.code + ": " + error.message;

        console.error(error);
    });
});


async function findUser(email) {

    await Excel.run(async (context) => {

        const table =
            context.workbook.tables.getItem("FilteredTable");

        const body =
            table.getDataBodyRange();

        body.load("values");

        await context.sync();

        const rows = body.values;
        
        const matchingRow = rows.find(row => {


            const rowEmails = String(row[2] || "")
                .split(";")
                .map(e => e.trim().toLowerCase());

            return rowEmails.includes(email.trim().toLowerCase());
        });

        if (matchingRow) {

            document.getElementById("userEmail").textContent =
                `${matchingRow[0]} (${email})`;

            document.getElementById("status").textContent =
                "Voter identified";

            document.getElementById("votingCard").style.display = "block";

            document.getElementById("submitVote").onclick = () => {
                submitVote(email);
            };

        } else {

            document.getElementById("status").textContent =
                "You are not listed as a voter.";
        }
    });
}


async function sendToAlternate(email) {

    const alternateEmail =
        document.getElementById("alternateEmail").value.trim().toLowerCase();

    if (!alternateEmail) {
        document.getElementById("status").textContent =
            "Please enter an alternate voter's email.";
        return;
    }

    if (alternateEmail === email.trim().toLowerCase()) {
        document.getElementById("status").textContent =
            "Alternate voter must be different from you.";
        return;
    }

    document.getElementById("status").textContent =
        "Sending voting authority to alternate...";

    try {

        await Excel.run(async (context) => {
            const sheet = 
                context.workbook.worksheets.getItem("Filtered Output");

            sheet.protection.pauseProtection(SHEET_PASSWORD);

            try {

                const table =
                    context.workbook.tables.getItem("FilteredTable");

                const body =
                    table.getDataBodyRange();

                body.load("values");

                await context.sync();

                const rows = body.values;

                const rowIndex = rows.findIndex(row => {

                    const rowEmails = String(row[2] || "")
                        .split(";")
                        .map(e => e.trim().toLowerCase());

                    return rowEmails.includes(email.trim().toLowerCase());
                });

                if (rowIndex === -1) {

                    document.getElementById("status").textContent =
                        "Could not find your voter row.";

                    return;
                }

                const currentEmail = String(rows[rowIndex][2] || "")
                    .trim();

                // Add alternate after the existing email
                body.getCell(rowIndex, 2).values = [[
                    currentEmail + ";" + alternateEmail
                ]];

                await context.sync();

                // Alternate has been entered, so disable normal submission
                document.getElementById("submitVote").disabled = true;

                document.getElementById("sendtoalt").disabled = true;

            } finally {

                sheet.protection.resumeProtection();
                await context.sync();

            }

        });

        document.getElementById("status").textContent =
            "Alternate voter added.";

    } catch (error) {

        document.getElementById("status").textContent =
            "Unable to add alternate voter.";

        console.error(error);
    }
}



async function submitVote(email) {

    const selectedVote =
        document.querySelector('input[name="vote"]:checked');

    const comments =
        document.getElementById("comments").value;

    if (!selectedVote) {

        document.getElementById("status").textContent =
            "Please select Y, N, or A.";

        return;
    }

    const vote = selectedVote.value;

    document.getElementById("status").textContent =
        "Submitting vote...";

    try {

        await Excel.run(async (context) => {

            const sheet =
                context.workbook.worksheets.getItem("Filtered Output");

            sheet.protection.pauseProtection(SHEET_PASSWORD);

            try {

                const table =
                    context.workbook.tables.getItem("FilteredTable");

                const body =
                    table.getDataBodyRange();

                body.load("values");

                await context.sync();

                const rows = body.values;

                const rowIndex = rows.findIndex(row => {

                    const rowEmails = String(row[2] || "")
                        .split(";")
                        .map(e => e.trim().toLowerCase());

                    return rowEmails.includes(email.trim().toLowerCase());
                });

            if (rowIndex === -1) {

                document.getElementById("status").textContent =
                    "Could not find your row.";

                return;
            }

            // Column D = index 3
            // Column E = index 4

            body.getCell(rowIndex, 3).values = [[vote]];
            body.getCell(rowIndex, 4).values = [[comments]];

            await context.sync();


            console.log("Vote submitted:", vote);

            } finally {
                
                sheet.protection.resumeProtection();
                await context.sync();
            }

        });

        document.getElementById("status").textContent =
            "Vote submitted successfully.";

    } catch (error) {

        document.getElementById("status").textContent =
            "Vote submission failed.";

        console.error(error);
    }
}