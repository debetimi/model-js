'use strict';


    // Object for Formatting Numbers
    var NumberFormatter = {

        formats: ['_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
                    '_($* #,##0_);_(* (#,##0);_($* "-"_);_(@_)',
                    '#,##0.0%_);(#,##0.0%)',
                    '#,##0.0x',
                    'general'],

        getNextFormat: function (numberFormat) {

            var x = numberFormat.length;
            var y = numberFormat[0].length;
            
            var allFormatsMatch = true;
            var nextFormat = numberFormat[0][0];

            for (var i = 0; i < x; i++) {
                for (var j = 0; j < y; j++) {
                    if (nextFormat !== numberFormat[i][j]) {
                        allFormatsMatch = false;
                        break;
                    }
                }
            }

            if (!allFormatsMatch) {
                return nextFormat;
            } else {
                var index = this.formats.indexOf(nextFormat);
                if (index === -1) {
                    return this.formats[0];
                } else {
                    return this.formats[(index + 1) % this.formats.length];
                }
            }
        }
    };


    var ColorFormatter = {

        formatCell: function (cell, formula) {

            var hasExternalReference = /^\=.*xls[xm]?\].*!/;
            var hasLocalReference = /^\=.*[a-zA-Z]+[0-9]+/;
            var hasForeignReference = /^\=.*!\$?[a-zA-Z]+\$?[0-9]+/;
            var hasConstantFormula = /^\=/;

            if (formula === "") {
                // empty string
                return; 
            } 

            if (isNaN(formula)) {

                // If cell contains reference outside of file
                if (hasExternalReference.test(formula)) {
                    cell.format.font.color = "red";
                }
                // If cell contains reference to different sheet
                else if (hasForeignReference.test(formula)) {
                    cell.format.font.color = "green";
                }

                // If cell contains local reference formula
                else if (hasLocalReference.test(formula)) {
                    cell.format.font.color = "black";
                }

                // If cell contains a constant formula 
                // e.g. =1 + SUM(10, 30)
                else if (hasConstantFormula.test(formula)) {
                    cell.format.font.color = "blue";
                }
            } else {
                // If here, cell is a hardcoded number
                cell.format.font.color ="blue";
            }
        },

        // Calls formatCell on all elements in the range 
        processRange: function (range) {
            var formulas = range.formulas;
            var x = formulas.length;
            var y = formulas[0].length;

            for (var i = 0; i < x; i++) {
                for (var j = 0; j < y; j++) {
                    this.formatCell(range.getCell(i,j), formulas[i][j]);
                }
            }
        }
    }

    // Fixes color formatting on current spreadsheet
    function colorize() {
        var t0 = performance.now();
        Excel.run(function (context) {
            var usedRange = context.workbook.getSelectedRange()
                .worksheet
                .getUsedRange();

            usedRange.load('formulas');

            return context.sync(usedRange)
                .then(function (ranged) {
                    ColorFormatter.processRange(ranged);
                })
                .then(function() {
                    var t1 = performance.now();
                    console.log("Runtime: " + (t1 - t0) + "ms");
                });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function toggleNumberFormat() {
        Excel.run(function (context) {
            var range = context.workbook.getSelectedRange();
            range.load('numberFormat');
            return context.sync(range).then(function(range) {
                range.numberFormat = NumberFormatter.getNextFormat(range.numberFormat);
            });
        });
    }

(function () {
    //Assign function to DOM element
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(colorize);
            $('#toggle-format').click(toggleNumberFormat);
        });
    };
})();
