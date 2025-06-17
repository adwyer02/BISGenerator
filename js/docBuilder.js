import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    VerticalAlign,
    TabStopType
} from "../modules/docx.mjs";

let fileLoadBar = document.getElementById('fileLoadBar');
let downloadButton = document.getElementById('downloadButton');
let uploadText = document.getElementById('uploadText');
export let error = '';
export let warning = ''

export function buildDocx(docData, passWarning, passError) {
    error = passError;
    warning = passWarning;

    if(error === '') {
        const bisDoc = new Document({
            styles: {
                paragraphStyles: [
                    {
                        id: "12ptBold",
                        name: "12pt Bold",
                        basedOn: "Normal",
                        next: "Normal",
                        run: {
                            bold: true,
                            size: "12pt"
                        }
                    },
                    {
                        id: "14ptBold",
                        name: "14pt Bold",
                        basedOn: "Normal",
                        next: "Normal",
                        run: {
                            font: "Arial",
                            bold: true,
                            size: "14pt"
                        }
                    },
                    {
                        id: "14ptBoldHighlighted",
                        name: "14pt Bold Highlighted",
                        basedOn: "Normal",
                        next: "Normal",
                        run: {
                            font: "Arial",
                            highlight: 'yellow',
                            bold: true,
                            size: "14pt"
                        }
                    }
                ]
            },
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 180,
                            bottom: 180,
                            left: 360,
                            right: 360
                        }
                    }
                },
                children: [
                    buildHeading(docData),
                    buildTable(docData['table']),
                    new Paragraph({
                        text: "Additional Information: ",
                        style: "14ptBoldHighlighted"
                    }),
                    new Paragraph({
                        text: "Details & Consistency are primary concerns for clients. Ensuring that the building is 100% every night is priority.",
                        style: "14ptBold"
                    }),
                    new Paragraph({
                        text: "",
                        style: "14ptBold"
                    }),
                    new Paragraph({
                        text: "COMPLETED BY: ___________________________     DATE: _______",
                        style: "14ptBold"
                    }),
                    new Paragraph({
                        text: "",
                        style: "14ptBold"
                    }),
                    new Paragraph({
                        text: "INSPECTED BY: ____________________________     DATE: _______",
                        style: "14ptBold"
                    }),
                ],
            }],
        });
        const date = new Date();
        const year = `${date.getFullYear()}`;
        document.getElementById('downloadButton').addEventListener('click', function (e) {
            exportDocx(bisDoc, `${docData['buildingName']} BIS ${date.getMonth() + 1}.${date.getDate()}.${year.substring(2)}.docx`);
        })
    }
}

function buildHeading(docData) {
    return new Paragraph({
        text: `Building: ${docData['buildingName']}\tWeek Ending: __________________`,
        style: "14ptBold",
        tabStops: [{
            type: TabStopType.RIGHT,
            position: 11160
        }],
        spacing: {
            line: 300,
        }

    })
}

function buildTable(docData) {
    //get days/time frames to include
    let labelMap = {
        'monday': 'Mon',
        'tuesday': 'Tues',
        'wednesday': 'Wed',
        'thursday': 'Thurs',
        'friday': 'Fri',
        'saturday': 'Sat',
        'sunday': 'Sun',
        'weekly': 'W',
        'monthly': 'M',
        'quarterly': 'Q',
    }
    let includedColumns = [];

    let topRow = new TableRow({
        height: {
            value: 375,
        },
        children: [new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            margins: {
                left: 100
            },
            children: [new Paragraph({
                text: "Areas of Responsibility",
                style: "12ptBold"
            })]
        })]
    });

    Object.keys(docData).forEach(areaKey => {
        let area = docData[areaKey];
        Object.keys(area).forEach(sectionKey => {
            sectionKey.split(',').forEach(label => {
                if(!includedColumns.includes(labelMap[label])) {
                    includedColumns.push(labelMap[label]);
                }
            })
        })
    })

    //order labels
    let boxTitles = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'weekly', 'monthly', 'quarterly'];
    let fullWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    let sortedArray = [];
    boxTitles.forEach(title => {
        if(includedColumns.includes(labelMap[title])) {
            sortedArray.push(labelMap[title]);
        }
    })
    includedColumns = sortedArray;

    includedColumns.forEach(label => {
        let cell = new TableCell({
            verticalAlign: VerticalAlign.CENTER,
            width: {
                size: 560,
                type: WidthType.DXA
            },
            children: [
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun(label)],
                })
            ]
        })
        topRow.root.push(cell);
    })

    //create table
    const table = new Table({
        rows: [topRow],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
        borders: {
            top: {
                style: BorderStyle.DOUBLE
            },
            bottom: {
                style: BorderStyle.DOUBLE
            },
            left: {
                style: BorderStyle.DOUBLE
            },
            right: {
                style: BorderStyle.DOUBLE
            }
        }
    });

    //build table
    //get into areas
    Object.keys(docData).forEach(areaKey => {
        let area = docData[areaKey];
        let areaTitle = areaKey + ' (';
        let areaLabels = [];

        //build area title
        Object.keys(area).forEach(sectionKey => {
            sectionKey.split(',').forEach(label => {
                if(!areaLabels.includes(label)) {
                    areaLabels.push(label);
                }
            })
        });
        //sort labels
        let sortedArray2 = [];
        boxTitles.forEach(label => {
            if(areaLabels.includes(label)) {
                sortedArray2.push(label);
            }
        })
        areaLabels = sortedArray2;
        //append labels
        let appendIndex = 0;
        let thru = [];
        while(appendIndex < areaLabels.length) {
            let thisLabel = areaLabels[appendIndex];

            //thru
            if(areaLabels[appendIndex + 1] === fullWeek[fullWeek.indexOf(thisLabel) + 1] && areaLabels[appendIndex + 2] === fullWeek[fullWeek.indexOf(thisLabel) + 2]) {
                thru[0] = labelMap[thisLabel];
                thru[1] = labelMap[areaLabels[appendIndex + 2]];
                let addIndex = 3;
                if(areaLabels[appendIndex + 3] === fullWeek[fullWeek.indexOf(thisLabel) + 3]) {
                    thru[1] = labelMap[areaLabels[appendIndex + 3]];
                    addIndex++;
                    if(areaLabels[appendIndex + 4] === fullWeek[fullWeek.indexOf(thisLabel) + 4]) {
                        thru[1] = labelMap[areaLabels[appendIndex + 4]];
                        addIndex++;
                        if(areaLabels[appendIndex + 5] === fullWeek[fullWeek.indexOf(thisLabel) + 5]) {
                            thru[1] = labelMap[areaLabels[appendIndex + 5]];
                            addIndex++;
                            if(areaLabels[appendIndex + 6] === fullWeek[fullWeek.indexOf(thisLabel) + 6]) {
                                thru[1] = labelMap[areaLabels[appendIndex + 6]];
                                addIndex++;
                            }
                        }
                    }
                }
                appendIndex += addIndex;
                areaTitle += thru[0] + '-' + thru[1];
                if(appendIndex < areaLabels.length) {
                    areaTitle += ', '
                }
            } else if(appendIndex < areaLabels.length - 1) {
                areaTitle += labelMap[thisLabel] + ', ';
                appendIndex++;
            } else {
                areaTitle += labelMap[thisLabel];
                appendIndex++;
            }
        }
        areaTitle += ')';

        let areaRow = new TableRow({
            height: {
                value: 350,
            },
            children: [
                new TableCell({
                    verticalAlign: VerticalAlign.CENTER,
                    margins: {
                        left: 100
                    },
                    shading: {
                        fill: "A6A6A6"
                    },
                    children: [
                        new Paragraph({
                            children: [new TextRun({
                                text: areaTitle,
                                bold: true
                            })]
                        })
                    ]
                })
            ]
        });
        includedColumns.forEach(label => {
            areaRow.root.push(
                new TableCell({
                    verticalAlign: VerticalAlign.CENTER,
                    shading: {
                        fill: "A6A6A6"
                    },
                    width: {
                        size: 560,
                        type: WidthType.DXA
                    },
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun(label)],
                        })
                    ]
                })
            );
        })
        table.root.push(areaRow)

        //get into sections
        Object.keys(area).forEach(sectionKey => {
            let sectionLabels = sectionKey.split(',');
            for(let i = 0; i < sectionLabels.length; i++) {
                if(sectionLabels[i] === "") {
                    warning = `One or more sections did not include any of the available time slot keywords. Please either:<div style="color: black;">1) Review documentation for available keywords, correct the service agreement, and resubmit.<br>2) Download BIS and review with caution before using</div>`;
                } else {
                    sectionLabels[i] = labelMap[sectionLabels[i]];
                }
            }
            let section = docData[areaKey][sectionKey];

            //get into line items
            Object.keys(section).forEach(lineItemKey => {
                let lineItem = docData[areaKey][sectionKey][lineItemKey];
                let rowChildren = [new TableCell({
                    margins: {
                        left: 100,
                        top: 25,
                        bottom: 50
                    },
                    children: [new Paragraph(lineItem)]
                })];
                includedColumns.forEach(includedColumn => {
                    if(sectionLabels.includes(includedColumn)) {
                        rowChildren.push(
                            new TableCell({
                                shading: {
                                    fill: "FFFFFF"
                                },
                                children: [new Paragraph('')]
                            })
                        );
                    } else {
                        rowChildren.push(
                            new TableCell({
                                shading: {
                                    fill: "000000"
                                },
                                children: [new Paragraph('')]
                            })
                        );
                    }
                });
                table.root.push(
                    new TableRow({
                        children: rowChildren,
                    })
                )
            })
        })
    })

    return table;
}

function exportDocx(doc, fileName) {
    Packer.toBlob(doc).then((blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        a.click();
        URL.revokeObjectURL(url);
    });
}