'use strict';


    // Object for Formatting Numbers
    let NumberFormatter = function () {

        this.formats = ['_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
                '_($* #,##0_);_(* (#,##0);_($* "-"_);_(@_)',
                    '#,##0.0%_);(#,##0.0%)',
                '#,##0.0x',
                'general'];
    };

    NumberFormatter.prototype.getNextFormat = function (numberFormat) {

        let allFormatsMatch = true;
        const firstFormat = numberFormat[0][0];

        for (let i = 0; i < numberFormat.length; i++) {
            for (let j = 0; j < numberFormat[0].length; j++) {
                if (firstFormat !== numberFormat[i][j]) {
                    allFormatsMatch = false;
                    break;
                }
            }
        }

        if (!allFormatsMatch) {
            return firstFormat;
        } else {
            let index = this.formats.indexOf(firstFormat);
            return this.formats[(index + 1) % this.formats.length];
        }
    };


    var ColorFormatter = {

        regExpToColor: [[/^\=.*xls[xm]?\].*!/, "red"],
                        [/^\=.*\$?[a-zA-Z]+\$?[0-9]+/, "green"],
                        [/^\=.*!\$?[a-zA-Z]+\$?[0-9]+/, "black"],
                        [/^\=/, "blue"]],

        formatCell: function (cell, formula) {

            if (formula === "") {
                // empty string
                return; 
            } 

            // Cell is a number
            if (!isNaN(formula)) {
                cell.format.font.color = "blue";
                return;
            }
            
            for (let i = 0; i < this.regExpToColor.length; i++) {
                if (this.regExpToColor[i][0].test(formula)) {
                    cell.format.font.color = this.regExpToColor[i][1];
                    return;
                }
            }
        },

        // Calls formatCell on all elements in the range 
        processRange: function (range) {
            var formulas = range.formulas;
            for (let i = 0; i < formula.length; i++) {
                for (let j = 0; j < formulas[0].length; j++) {
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
                    let t1 = performance.now();
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
            range.load("numberFormat");
            return context.sync(range).then(function(range) {
                range.numberFormat = new NumberFormatter().getNextFormat(range.numberFormat);
            });
        });
    }

(function () {
    //Assign function to DOM element
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("#set-color").click(colorize);
            $("#toggle-format").click(toggleNumberFormat);
        });
    };
})();
