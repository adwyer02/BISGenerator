import mammoth from '../modules/mammoth.module.js';
import {buildDocx, error, warning} from "./docBuilder.js";

document.getElementById("upload").addEventListener("change", function(event) {
    document.getElementById("mainContent").style.height = '830px';

    let fileLoadBar = document.getElementById('fileLoadBar');
    let downloadButton = document.getElementById('downloadButton');
    let uploadText = document.getElementById('uploadText');

    fileLoadBar.style.display = 'inline';
    downloadButton.style.display = 'none';
    uploadText.innerHTML = '';

    let boxTitles = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'weekly', 'monthly', 'quarterly'];
    let fullWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    let weekDays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
    let passWarning = '';
    let passError = '';

    const file = event.target.files[0];
    const fileName = file.name;

    //check for errors on initial file upload
    if(!(fileName.substring(fileName.length - 5, fileName.length) === '.docx') || !file) {
        if(fileName.substring(fileName.length - 4, fileName.length) === '.doc') {
            passError = `You have uploaded a file of type ".doc". This tool requires files of type ".docx". Visit <a href="https://cloudconvert.com/doc-to-docx" target="_blank">this website</a> to convert your ".doc" to ".docx", then try again.`;
        } else {
            passError = `Invalid file submission. Please upload a file of type ".docx".`;
        }
    }

    function showError(str) {
        downloadButton.style.display = 'none';
        fileLoadBar.style.display = 'none';
        uploadText.innerHTML = "ERROR: " + str;
        uploadText.classList.add('error');
    }
    function showWarning(str) {
        downloadButton.style.display = 'inline';
        fileLoadBar.style.display = 'none';
        uploadText.innerHTML = "WARNING: " + str;
        uploadText.classList.add('error');
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const arrayBuffer = reader.result;
        let docText = "";
        let dailyDays = [];

        if(passError !== '') {
            showError(passError);
        } else {
            mammoth.convertToHtml({arrayBuffer: arrayBuffer})
                .then(function (result) {
                    docText = result.value;

                    //format docText
                    docText = docText.replaceAll('&amp;', '&');
                    let runAnchor = true;
                    let searchIndexAnchor = 0;
                    while(runAnchor) {
                        let anchorStartIndex = docText.indexOf('<a', searchIndexAnchor)
                        if(anchorStartIndex > -1) {
                            let anchorEndIndex = docText.indexOf('</a>', anchorStartIndex) + 4;
                            let remString = docText.substring(anchorStartIndex, anchorEndIndex);
                            console.log(remString);
                            docText = docText.replace(remString, '');
                        } else {
                            runAnchor = false;
                        }
                    }

                    //error checks
                    if (!docText.substring(docText.indexOf('>'), docText.indexOf('<', docText.indexOf('>'))).toLowerCase().includes('service agreement') && !fileName.toLowerCase().includes('service agreement')) {
                        showError('Uploaded document is not a service agreement.');
                    } else {
                        if (!(docText.includes('<h1>') && docText.includes('<h2>') && docText.includes('<h3>') && docText.includes('<em>') && docText.includes('<ul>') && docText.includes('<li>'))) {
                            showError('Uploaded document does not follow formatting rules. Please review documentation on how to properly format the document.')
                        } else {
                            uploadText.classList.remove('error');
                            let progress = 0;
                            incrementProgress()

                            function incrementProgress() {
                                if (progress < 100) {
                                    progress += 10;
                                    fileLoadBar.value = progress;
                                    setTimeout(incrementProgress, 20);
                                } else {
                                    if (error !== '') {
                                        showError(error);
                                    } else if (warning !== '') {
                                        showWarning(warning);
                                    } else {
                                        uploadText.innerHTML = `File uploaded successfully: <img src="../imgs/docx_icon.png" style="width:20px; vertical-align: middle;" alt="docx icon"><a href="${URL.createObjectURL(file)}" download="${fileName}"> ${fileName}</a>`;
                                        downloadButton.style.display = 'inline';
                                        fileLoadBar.style.display = 'none';
                                    }
                                }
                            }
                        }

                        let docData = getDocData(docText);
                        console.log(docData);
                        buildDocx(docData, passWarning, passError);
                    }

                    function getDocData() {
                        let retMap = {};
                        retMap['table'] = {};

                        //get building name
                        let nameStartIndex = docText.indexOf("<h1>", 0) + 4;
                        retMap["buildingName"] = docText.substring(nameStartIndex, docText.indexOf("</h1>", nameStartIndex)).trim();

                        //get daily days
                        let run1 = true;
                        let startIndex1 = 0;

                        while (run1) {
                            let titleStartIndex = docText.indexOf("<h2>", startIndex1);

                            if (titleStartIndex === -1) {
                                run1 = false;
                                break;
                            }

                            startIndex1 = titleStartIndex + 1;
                            let titleEndIndex = docText.indexOf("</h2>", startIndex1);
                            let title = docText.substring(titleStartIndex + 4, titleEndIndex);

                            getSections(title);
                        }

                        //get building data
                        let run2 = true;
                        let startIndex2 = 0;

                        while (run2) {
                            let titleStartIndex = docText.indexOf("<h2>", startIndex2);

                            if (titleStartIndex === -1) {
                                run2 = false;
                                break;
                            }

                            startIndex2 = titleStartIndex + 1;
                            let titleEndIndex = docText.indexOf("</h2>", startIndex2);
                            let title = docText.substring(titleStartIndex + 4, titleEndIndex);

                            retMap['table'][title] = getSections(title);
                        }
                        return retMap;
                    }

                    function getSections(title) {
                        let retObj = {};
                        let startIndex = docText.indexOf(title);
                        startIndex += title.length;

                        let sectionEndIndex = -1;
                        if (docText.indexOf('<h2>', startIndex) === -1) {
                            sectionEndIndex = docText.indexOf('<h3>', startIndex);
                        } else {
                            sectionEndIndex = docText.indexOf('<h2>', startIndex);
                        }

                        let run = true;
                        if (sectionEndIndex === -1) {
                            run = false;
                        }
                        while (run) {
                            let sectionTitleStartIndex = docText.indexOf("<em>", startIndex);

                            if (sectionTitleStartIndex === -1 || sectionTitleStartIndex >= sectionEndIndex) {
                                run = false;
                                break;
                            }

                            let sectionTitleEndIndex = docText.indexOf("</em>", startIndex);
                            let sectionTitle = docText.substring(sectionTitleStartIndex, sectionTitleEndIndex).toLowerCase();


                            let sectionBoxes = [];

                            //check for daily
                            let dailySynonyms = ['daily', 'every day', 'nightly', 'every night'];
                            dailySynonyms.forEach(str => {
                                if (sectionTitle.includes(str)) {
                                    dailyDays.forEach(day => {
                                        if (!sectionBoxes.includes(day)) {
                                            sectionBoxes.push(day);
                                        }
                                    });
                                }
                            })

                            //check for weekdays
                            if (sectionTitle.includes('weekday')) {
                                weekDays.forEach(day => {
                                    if (!sectionBoxes.includes(day)) {
                                        sectionBoxes.push(day);
                                    }
                                    if (!dailyDays.includes(day)) {
                                        dailyDays.push(day);
                                    }
                                })
                            }

                            //check for seven days a week
                            let sevenSynonyms = ['seven days a week', 'seven days per week', 'seven nights a week', 'seven nights per week', 'seven days/week', 'seven nights/week', 'Every day of the week'];
                            sevenSynonyms.forEach(str => {
                                if (sectionTitle.includes(str)) {
                                    fullWeek.forEach(day => {
                                        if (!sectionBoxes.includes(day)) {
                                            sectionBoxes.push(day);
                                        }
                                        if (!dailyDays.includes(day)) {
                                            dailyDays.push(day);
                                        }
                                    })
                                }
                            });

                            //check for dash and through (monday-friday)
                            checkRange("-");
                            checkRange("through");

                            function checkRange(str) {
                                let runRanges = true;
                                let rangeStartIndex = 0;
                                while (runRanges) {
                                    if (sectionTitle.substring(rangeStartIndex).includes(str) && rangeStartIndex < sectionTitle.length) {
                                        let dashIndex = sectionTitle.indexOf(str);
                                        let rangeStart = 0;
                                        let rangeEnd = 0;

                                        for (let i = 0; i < fullWeek.length; i++) {
                                            if (sectionTitle.substring(dashIndex - 11, dashIndex).includes(fullWeek[i])) {
                                                rangeStart = i;
                                            }
                                            if (sectionTitle.substring(dashIndex, dashIndex + str.length + 11).includes(fullWeek[i])) {
                                                rangeEnd = i;
                                            }
                                        }

                                        let rangeDays = [];
                                        if (rangeStart <= rangeEnd) {
                                            rangeDays = fullWeek.slice(rangeStart, rangeEnd);
                                        } else {
                                            rangeDays = fullWeek.slice(rangeEnd).concat(fullWeek.slice(0, rangeStart));
                                        }

                                        rangeDays.forEach(rangeDay => {
                                            if (!sectionBoxes.includes(rangeDay)) {
                                                sectionBoxes.push(rangeDay);
                                            }
                                            if (!dailyDays.includes(rangeDay)) {
                                                dailyDays.push(rangeDay);
                                            }
                                        });

                                        rangeStartIndex = dashIndex + str.length;
                                    } else {
                                        runRanges = false;
                                    }
                                }
                            }

                            //check for individual labels
                            boxTitles.forEach(boxTitle => {
                                if (sectionTitle.includes(boxTitle)) {
                                    if (!sectionBoxes.includes(boxTitle)) {
                                        sectionBoxes.push(boxTitle);
                                    }
                                    if (!dailyDays.includes(boxTitle) && fullWeek.includes(boxTitle)) {
                                        dailyDays.push(boxTitle);
                                    }
                                }
                            });

                            startIndex = sectionTitleEndIndex + 1;

                            retObj[sectionBoxes.toString()] = getBullets(sectionTitleEndIndex);
                        }

                        return retObj;
                    }

                    function getBullets(startIndex) {
                        let retList = [];

                        if (docText.indexOf('<em>', startIndex) < docText.indexOf('<ul>', startIndex)) {
                            passWarning = `Multiple section titles found for one section, document is likely too complex for a BIS to be automatically generated.<br>If you choose to download BIS anyway, please review with caution and correct any mistakes.`;
                        } else {
                            let listStartIndex = docText.indexOf("<ul>", startIndex);
                            let listEndIndex = docText.indexOf("</ul>", startIndex);
                            startIndex = listStartIndex;

                            let run = true;
                            while (run) {
                                let itemStartIndex = docText.indexOf("<li>", startIndex);

                                if (itemStartIndex === -1 || itemStartIndex >= listEndIndex) {
                                    run = false;
                                    break;
                                }

                                startIndex = itemStartIndex + 1;
                                let itemEndIndex = docText.indexOf("</li>", startIndex);
                                let item = docText.substring(itemStartIndex + 4, itemEndIndex);
                                retList.push(item);
                            }
                        }

                        return retList;
                    }
                })
                .catch(function (err) {
                    console.error("Error:", err);
                });
        }
    };
    reader.readAsArrayBuffer(file);
});