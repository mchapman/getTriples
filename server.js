var XLS = require('xlsjs')
    , path = require('path');

var config = {
    NUMBERS_REQ_FOR_TAKE : 3
};

if (process.argv.length >= 3) {
    processFile(process.argv[2]);
}

function processFile(filename) {
    var type = path.extname(filename);
    switch (type) {
        case ".xls" :
            handleXls(filename);
            break;
        default: console.log("Unsupported file type in "+filename);
    }
}

function handleXls(filename) {
    var xls = XLS.readFile(filename);
    for (var i=0; i < xls.SheetNames.length; i++) {
        handleXlsSheet(xls.SheetNames[i] , xls.Sheets[xls.SheetNames[i]]);
    }
}

function handleRowCategory(data, col, row, rowCategories) {
    // figure out what row we get the category from
    for (var backCol = col - 1; backCol >= 0; backCol --) {
        var backColRef = String.fromCharCode(backCol+65)+(row+1)
            , categoryCell = data[backColRef];

        if (categoryCell && categoryCell.t === 's') {
            rowCategories[row] = categoryCell.v;
            break;
        }
    }
    return rowCategories[row];
}

function handleXlsSheet(name, data) {
    if (data['!mergeCells']) {
        // Copy data in merges into all the merged cells so we can easily find it when looking for column headings
        for (var mergeCell = 0; mergeCell < data['!mergeCells'].length; mergeCell++) {
            var mergeData = data['!mergeCells'][mergeCell]
                , src = data[XLS.utils.encode_cell(mergeData.s)]
                , destRef;
            for (var mCol = mergeData.s.c; mCol <= mergeData.e.c; mCol++) {
                for (var mRow = mergeData.s.r; mRow <= mergeData.e.r; mRow++) {
                    destRef = XLS.utils.encode_cell({c:mCol, r:mRow});
                    data[destRef] = src;
                }
            }
        }
    }

    var rowCategories = [];
    // Look for columns of numbers in the data
    for (var col=data['!range'].s.c; col <= data['!range'].e.c; col++) {
        var numberCount = 0,
            numberCol = false,
            columnHeading = undefined;
        for (var row=data['!range'].s.r; row <= data['!range'].e.r; row++) {
            var cellRef = String.fromCharCode(col+65)+(row+1)
                , cell = data[cellRef];
            if (cell) {
                if (cell.t === 'n') {
                    if (!numberCol) {
                        numberCount++;
                        if (numberCount >= config.NUMBERS_REQ_FOR_TAKE) {
                            // Get what looks like a column heading
                            var colHeadCells = [];

                            for (var backRow = row - config.NUMBERS_REQ_FOR_TAKE; backRow >= 0; backRow --) {
                                var backRowRef = String.fromCharCode(col+65)+(backRow+1)
                                    , textCell = data[backRowRef];

                                if (textCell && textCell.t === 's') {
                                    colHeadCells.unshift(textCell.v);
                                } else {
                                    columnHeading = colHeadCells.join(' - ');
                                    break;
                                }
                            }

                            if (colHeadCells.length > 0) {
                                numberCol = true;
                                // We have something that looks like data - loop back and get the bits we skipped
                                // while we were making our minds up

                                for (var numberRow = row - config.NUMBERS_REQ_FOR_TAKE + 1; numberRow <= row; numberRow++ ) {
                                    var numberCellRef = String.fromCharCode(col+65)+(numberRow+1);
                                    console.log(numberCellRef, columnHeading, rowCategories[numberRow] || handleRowCategory(data, col, numberRow, rowCategories) , data[numberCellRef].v)
                                }
                            }
                        }
                    } else {
                        var triple = {heading:columnHeading, category: rowCategories[row]|| handleRowCategory(data, col, row, rowCategories) , value: data[cellRef].v};
                        console.log(cellRef, triple)
                    }
                } else {
                    numberCount = 0;
                }

            } else {
                numberCount = 0;
            }
        }
    }
}
