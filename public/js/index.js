const models = window["powerbi-client"].models;
const reportContainer = document.getElementById("report-container"),
    visualContainer = document.getElementById("visual-container");

visualContainer.style.height = "500px";

// Set to true to embed only the visual. 
// Note that you need to set it with your visual's metadata for it to work.
const showVisualEmbed = true;

// Initialize iframe for embedding report
powerbi.bootstrap(reportContainer, { type: "report" });
// Initialize iframe for embedding visual
if (showVisualEmbed)
    powerbi.bootstrap(visualContainer, { type: "visual" });

// Request to get the report details from the API and pass it to the UI
const xhr = new XMLHttpRequest;
xhr.open("GET", "/getEmbedToken", true);
xhr.setRequestHeader("Content-Type", "application/json");
xhr.onload = async function () {
    if (xhr.status >= 200 && xhr.status < 300) {
        const embedData = JSON.parse(xhr.responseText);

        // Create a config object with type of the object, Embed details and Token Type
        let reportLoadConfig = {
            type: "report",
            tokenType: models.TokenType.Embed,
            accessToken: embedData.accessToken,

            // Use other embed report config based on the requirement. We have used the first one for demo purpose
            embedUrl: embedData.embedUrl[0].embedUrl,

            // Enable this setting to remove gray shoulders from embedded report
            settings: {
                navContentPaneEnabled: false
            }
        };
        // You must change the name of the page and visual according to your report.
        const visualPageName = "ReportSection2e5116d592f50302d0cc",
            visualName = "3e29ffe001b19c75503e";
        let visualLoadConfig = {
            ...reportLoadConfig,
            type: "visual",
            pageName: visualPageName,
            visualName,
            settings: {
                "filterPaneEnabled": false,
                "navContentPaneEnabled": false,
                "layoutType": 1,
                "customLayout": {
                    "displayOption": 0,
                    "pageSize": {
                        "type": 4,
                        "width": 862,
                        "height": 500
                    },
                    "pagesLayout": {
                        [visualPageName]: {
                            "defaultLayout": {
                                "displayState": {
                                    "mode": 1
                                }
                            },
                            "visualsLayout": {
                                [visualName]: {
                                    "displayState": {
                                        "mode": 0
                                    },
                                    "x": 1,
                                    "y": 1,
                                    "z": 1,
                                    "width": 862,
                                    "height": 500
                                }
                            }
                        }
                    }
                }
            }
        };

        // Use the token expiry to regenerate Embed token for seamless end user experience
        tokenExpiry = embedData.expiry;

        // Embed Power BI report when Access token and Embed URL are available
        let report = powerbi.embed(reportContainer, reportLoadConfig);

        // Clear any other loaded handler events
        report.off("loaded");

        // Triggers when a schema is successfully loaded
        report.on("loaded", function () {
            console.log("Report load successful");
        });

        // Clear any other rendered handler events
        report.off("rendered");

        // Triggers when a report is successfully embedded in UI
        report.on("rendered", async function () {
            console.log("Report render successful");

            const pages = await report.getPages();
            pages.forEach(async (page) => {
                console.log(`\n---------Start page name: ${page.name} ---------`)
                const visuals = await page.getVisuals();
                visuals.forEach(visual => {
                    console.log({
                        name: visual.name,
                        type: visual.type,
                        title: visual.title,
                        layout: visual.layout
                    })
                })
                console.log(`\n---------End page name: ${page.name} ---------`)
            });

        }); 

        // Clear any other error handler events
        report.off("error");

        // Handle embed errors
        report.on("error", function (event) {
            let errorMsg = event.detail;
            console.error('Report error:', errorMsg);
            return;
        });

        if (showVisualEmbed) {
            let visual = powerbi.embed(visualContainer, visualLoadConfig);
            visual.off("loaded");
            visual.on("loaded", function () {
                console.log("Visual load successful");
            });
            visual.off("rendered");
            visual.on("rendered", async function () {
                console.log("Visual render successful");
                console.log('visual', visual)
            });
            visual.off("error");
            visual.on("error", function (event) {
                let errorMsg = event.detail;
                console.error('Visual error:', errorMsg);
                return;
            });
        } else {
            visualContainer.innerHTML = 'You must set `showVisualEmbed` to `true` at `public/js/index.js` and and configure the metadata of the visual to be embedded. More details in the article mentioned above.'
        }

    } else {
        errorHandler(xhr);
    }
};

xhr.onerror = errorHandler;

function errorHandler(err) {
    let errContainer = document.querySelector(".error-container");
    document.querySelectorAll(".embed-container").forEach(el => el.classList.add("d-none"));

    // Get the error message from err object
    let errMsg = JSON.parse(err.responseText)['error'];

    // Split the message with \r\n delimiter to get the errors from the error message
    let errorLines = errMsg.split("\r\n");

    // Create error header
    let errHeader = document.createElement("p");
    let strong = document.createElement("strong");
    let node = document.createTextNode("Error Details:");

    // Add the error header in the container
    strong.appendChild(node);
    errHeader.appendChild(strong);
    errContainer.appendChild(errHeader);

    // Create <p> as per the length of the array and append them to the container
    errorLines.forEach(element => {
        let errorContent = document.createElement("p");
        let node = document.createTextNode(element);
        errorContent.appendChild(node);
        errContainer.appendChild(errorContent);
    });
}

xhr.send();