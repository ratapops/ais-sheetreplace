const express = require("express");
const { google } = require("googleapis");

const spreadsheetId = "1qiS3fYu4sZcVKB4jLHZ-JUNUNfADwPBoyd4CGNB8kLQ";
const port = 1111;

const users = [
    {
        keyFile: "gg-sheet01.json",
        rows: 0,
    },
    {
        keyFile: "gg-sheet02.json",
        rows: 0,
    },
    {
        keyFile: "gg-sheet03.json",
        rows: 0,
    },
    {
        keyFile: "gg-sheet04.json",
        rows: 0,
    },
    {
        keyFile: "gg-sheet05.json",
        rows: 0,
    },
];

const app = express();

app.get("/replaceGuidance", async (req, res) => {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: "gg-sheet01.json",
            scopes: "https://www.googleapis.com/auth/spreadsheets",
        });

        const client = await auth.getClient();

        const sheets = google.sheets({ version: "v4", auth: client });

        const guidanceVal = await sheets.spreadsheets.values.get({
            spreadsheetId,
            auth,
            range: "Test!A:A",
        });

        const linkSheetVal = await sheets.spreadsheets.values.get({
            spreadsheetId,
            auth,
            range: "Link!A:B",
        });
        const guidanceValLength = guidanceVal.data.values.length;
        const linkSheetValLength = linkSheetVal.data.values.length;

        const accValArr = [];
        const ikmValArr = [];
        const TestLinkArr = [];

        if (linkSheetValLength >= 1) {
            linkSheetVal.data.values.map((arrVal) => {
                accValArr.push(arrVal[0]);
                ikmValArr.push(arrVal[1]);
            });
        }
        if (guidanceValLength >= 1) {
            guidanceVal.data.values.map((arrVal) => {
                TestLinkArr.push(arrVal[0]);
            });
        }
        let updatedArr = [];
        if (
            accValArr.length > 0 &&
            ikmValArr.length > 0 &&
            TestLinkArr.length > 0
        ) {
            updatedArr = TestLinkArr.map((str) => {
                for (let i = 0; i < accValArr.length; i++) {
                    str = str.replace(
                        new RegExp(accValArr[i], "g"),
                        ikmValArr[i]
                    );
                }
                return str;
            });

            if (updatedArr.length > 0) {
                const totalRows = updatedArr.length;
                const rowsPerUser = Math.round(totalRows / users.length);

                users.map((u) => {
                    u.rows = rowsPerUser;
                });

                await Promise.all(
                    users.map(async (user) => {
                        const auth = new google.auth.GoogleAuth({
                            keyFile: user.keyFile,
                            scopes: "https://www.googleapis.com/auth/spreadsheets",
                        });

                        const client = await auth.getClient();

                        const sheets = google.sheets({
                            version: "v4",
                            auth: client,
                        });

                        // Calculate range based on rows per user
                        const startRow = 1 + rowsPerUser * users.indexOf(user);
                        const endRow = startRow + user.rows - 1;

                        // Loop through each row in the range and update the value
                        for (let row = startRow; row <= endRow; row++) {
                            const rangeRow = row + 1;
                            const range = `Test!B${rangeRow}:B${rangeRow}`;

                            await sheets.spreadsheets.values.update({
                                auth,
                                spreadsheetId,
                                valueInputOption: "RAW",
                                range,
                                resource: {
                                    values: [[updatedArr[row]]],
                                },
                                // Delay before making next API request
                            });

                            await new Promise((resolve) =>
                                setTimeout(resolve, 700)
                            );

                            // console.log(
                            //     `Done updating sheet for User: ${user.keyFile}, Row: ${row}`
                            // );
                        }
                    })
                );
            }
        }

        res.status(200).send(updatedArr);
    } catch (err) {
        console.error("Error updating sheets:", err);
        res.status(500).send("Error updating sheets");
    }
});

app.get("/replaceKBLink", async (req, res) => {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: "gg-sheet01.json",
            scopes: "https://www.googleapis.com/auth/spreadsheets",
        });

        const client = await auth.getClient();

        const sheets = google.sheets({ version: "v4", auth: client });

        const sheet1Val = await sheets.spreadsheets.values.get({
            spreadsheetId,
            auth,
            range: "Test!A:A",
        });

        const linkSheetVal = await sheets.spreadsheets.values.get({
            spreadsheetId,
            auth,
            range: "Link!A:B",
        });

        const linkValLength = linkSheetVal.data.values.length;
        const sheet1ValLength = sheet1Val.data.values.length;
        const accValArr = [];
        const ikmValArr = [];
        const TestLinkArr = [];

        if (linkValLength >= 1) {
            linkSheetVal.data.values.map((arrVal) => {
                accValArr.push(arrVal[0]);
                ikmValArr.push(arrVal[1]);
            });
        }
        if (sheet1ValLength >= 1) {
            sheet1Val.data.values.map((arrVal) => {
                TestLinkArr.push(arrVal[0]);
            });
        }
        let updatedArr = [];
        if (
            accValArr.length > 0 &&
            ikmValArr.length > 0 &&
            TestLinkArr.length > 0
        ) {
            updatedArr = TestLinkArr.map((item) => {
                if (item != undefined || item != null) {
                    const index = accValArr.indexOf(item);
                    if (index !== -1) {
                        return ikmValArr[index];
                    }
                }
                return item;
            });
        }

        if (updatedArr.length > 0) {
            const totalRows = updatedArr.length;
            const rowsPerUser = Math.round(totalRows / users.length);

            users.map((u) => {
                u.rows = rowsPerUser;
            });

            await Promise.all(
                users.map(async (user) => {
                    const auth = new google.auth.GoogleAuth({
                        keyFile: user.keyFile,
                        scopes: "https://www.googleapis.com/auth/spreadsheets",
                    });

                    const client = await auth.getClient();

                    const sheets = google.sheets({
                        version: "v4",
                        auth: client,
                    });

                    // Calculate range based on rows per user
                    const startRow = 1 + rowsPerUser * users.indexOf(user);
                    const endRow = startRow + user.rows - 1;

                    // Loop through each row in the range and update the value
                    for (let row = startRow; row <= endRow; row++) {
                        const rangeRow = row + 1;
                        const range = `Test!B${rangeRow}:B${rangeRow}`;

                        await sheets.spreadsheets.values.update({
                            auth,
                            spreadsheetId,
                            valueInputOption: "RAW",
                            range,
                            resource: {
                                values: [[updatedArr[row]]],
                            },
                            // Delay before making next API request
                        });

                        await new Promise((resolve) =>
                            setTimeout(resolve, 100)
                        );

                        // console.log(
                        //     `Done updating sheet for User: ${user.keyFile}, Row: ${row}`
                        // );
                    }
                })
            );
        }

        // res.status(200).send(updatedArr);
        res.status(200).send("Done load /replaceKBLink");
    } catch (err) {
        console.error("Error updating sheets:", err);
        res.status(500).send("Error updating sheets");
    }
});

app.get("/replaceLink", async (req, res) => {
    const auth = new google.auth.GoogleAuth({
        keyFile: "gg-sheet01.json",
        scopes: "https://www.googleapis.com/auth/spreadsheets",
    });

    const client = await auth.getClient();

    const sheets = google.sheets({ version: "v4", auth: client });

    const getRows = await sheets.spreadsheets.values.get({
        spreadsheetId,
        auth,
        range: "Sheet1!K:M",
    });
    const rowsLength = getRows.data.values.length;
    if (rowsLength >= 1) {
        for (let i = 0; i < rowsLength; i++) {
            // skip header sheet
            if (i >= 1) {
                const element = getRows.data.values[i];
                const guidanceText = element[0];
                const guidanceLink = element[1];
                const newGuidanceLink = element[2];

                if (
                    guidanceLink !== undefined &&
                    newGuidanceLink !== undefined
                ) {
                    const guidanceLinkArr = guidanceLink.split(", ");
                    const newGuidanceLinkArr = newGuidanceLink.split(", ");

                    let replacedString = guidanceText;

                    if (guidanceLinkArr.length == newGuidanceLinkArr.length) {
                        for (let j = 0; j < guidanceLinkArr.length; j++) {
                            replacedString = replacedString.replace(
                                guidanceLinkArr[j],
                                newGuidanceLinkArr[j]
                            );
                        }
                    }

                    const rangeIndex = i + 1;
                    const updateRange =
                        "Sheet1!N" + rangeIndex + ":N" + rangeIndex;

                    await sheets.spreadsheets.values.update({
                        auth,
                        spreadsheetId,
                        valueInputOption: "RAW",
                        range: updateRange,
                        resource: {
                            values: [[replacedString]],
                        },
                    });
                    // Delay before making next API request
                    // await new Promise((resolve) => setTimeout(resolve, 1000)); // Delay for 1 second (1000 ms)
                }
            }
        }
    }

    res.send("Done load /replaceLink");
});

app.get("/getLink", async (req, res) => {
    try {
        // Calculate total number of rows
        const totalRows = 2000;

        // Calculate number of rows per user
        const rowsPerUser = totalRows / users.length;

        users.map((u) => {
            u.rows = rowsPerUser;
        });

        // Use Promise.all() to make API calls in parallel for all users
        await Promise.all(
            users.map(async (user) => {
                const auth = new google.auth.GoogleAuth({
                    keyFile: user.keyFile,
                    scopes: "https://www.googleapis.com/auth/spreadsheets",
                });

                const client = await auth.getClient();

                const sheets = google.sheets({ version: "v4", auth: client });

                const getRows = await sheets.spreadsheets.values.get({
                    spreadsheetId,
                    auth,
                    range: "Sheet1!K:K",
                });
                const rowsLength = getRows.data.values.length;

                if (rowsLength >= 1) {
                    // loop in to values in column
                    for (let i = 0; i < rowsLength; i++) {
                        // skip header sheet
                        if (i >= 1) {
                            const urls = [];
                            const element = getRows.data.values[i];

                            const urlPattern =
                                /((?:https?:\/\/|www\.)[\w-]+(\.[\w-]+)+([^\s()<>]+(?:\([\w\d]+\)|([^()\s]+))))/gi;

                            let match;
                            while (
                                (match = urlPattern.exec(element[0])) !== null
                            ) {
                                urls.push(match[1]);
                            }
                            // make unique array
                            const uniqueArray = urls.filter(
                                (value, index, self) => {
                                    return self.indexOf(value) === index;
                                }
                            );

                            const userIndex = users.indexOf(user); // 0 1 2 3 4
                            const rangeIndex = i + 1;

                            // Calculate range based on rows per user
                            const startRow =
                                rowsPerUser * users.indexOf(user) + 1;
                            const endRow = startRow + user.rows - 1;

                            const range = `Sheet1!L${rangeIndex}:L${rangeIndex}`;

                            const urlVal = uniqueArray.join(", ");
                            // console.log(range)

                            await sheets.spreadsheets.values.update({
                                auth,
                                spreadsheetId,
                                valueInputOption: "RAW",
                                range,
                                resource: {
                                    values: [[urlVal]],
                                },
                            });
                            // Delay before making next API request
                            await new Promise((resolve) =>
                                setTimeout(resolve, 800)
                            ); // Delay for 1 second (1000 ms)
                        }
                    }
                }
                console.log(`Done updating sheet for User: ${user.keyFile}`);
            })
        );

        res.status(200).send("Done load /getLink");
    } catch (err) {
        console.error("Error updating sheets:", err);
        res.status(500).send("Error updating sheets");
    }
});

app.get("/", (req, res) => {
    res.send("Hello world");
});

app.listen(port, (req, res) => {
    console.log("Running server on " + port);
});
