document.getElementById("upload").addEventListener("change", function(event) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const arrayBuffer = reader.result;
        let docText = "";

        mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
            .then(function(result) {
                docText = result.value;
                document.getElementById("output").innerHTML = docText;

                let docData = getDocData(docText);
                console.log(docData);

                function getDocData() {
                    let run = true;
                    let startIndex = 0;
                    let retMap = {};
                    while(run) {
                        let titleStartIndex = docText.indexOf("<h1>", startIndex);

                        if(titleStartIndex === -1) {
                            run = false;
                            break;
                        }

                        startIndex = titleStartIndex + 1;
                        let titleEndIndex = docText.indexOf("</h1>", startIndex);
                        let title = docText.substring(titleStartIndex + 4, titleEndIndex);

                        retMap[formatTitle(title)] = getSections(title);
                    }
                    return retMap;
                }

                function formatTitle(str) {
                    return str.replaceAll('&amp;', '&');
                }

                function getSections(title) {
                    let retObj = {};
                    let boxTitles = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'weekly', 'monthly', 'quarterly'];
                    let startIndex = docText.indexOf(title);
                    startIndex += title.length;

                    let sectionEndIndex = -1;
                    if(docText.indexOf('<h1>', startIndex) === -1) {
                        sectionEndIndex = docText.indexOf('<h2>', startIndex);
                    } else {
                        sectionEndIndex = docText.indexOf('<h1>', startIndex);
                    }

                    let run = true;
                    if(sectionEndIndex === -1) {
                        run = false;
                    }
                    while(run) {
                        let sectionTitleStartIndex = docText.indexOf("<em>", startIndex);

                        if(sectionTitleStartIndex === -1 || sectionTitleStartIndex >= sectionEndIndex) {
                            run = false;
                            break;
                        }

                        let sectionTitleEndIndex = docText.indexOf("</em>", startIndex);
                        let sectionTitle = docText.substring(sectionTitleStartIndex, sectionTitleEndIndex).toLowerCase();

                        let sectionBoxes = [];
                        boxTitles.forEach(boxTitle => {
                            if(sectionTitle.includes(boxTitle)) {
                                sectionBoxes.push(boxTitle);
                            }
                        });

                        startIndex = sectionTitleEndIndex + 1;

                        retObj[sectionBoxes.toString()] = getBullets(sectionTitleEndIndex);
                    }

                    return retObj;
                }

                function getBullets(startIndex) {
                    let retList = [];

                    let listStartIndex = docText.indexOf("<ul>", startIndex);
                    let listEndIndex = docText.indexOf("</ul>", startIndex);
                    startIndex = listStartIndex;

                    let run = true;
                    while(run) {
                        let itemStartIndex = docText.indexOf("<li>", startIndex);

                        if(itemStartIndex === -1 || itemStartIndex >= listEndIndex) {
                            run = false;
                            break;
                        }

                        startIndex = itemStartIndex + 1;
                        let itemEndIndex = docText.indexOf("</li>", startIndex);
                        let item = docText.substring(itemStartIndex + 4, itemEndIndex);
                        retList.push(item);
                    }

                    return retList;
                }

            })
            .catch(function(err) {
                console.error("Error:", err);
            });
    };
    reader.readAsArrayBuffer(event.target.files[0]);
});