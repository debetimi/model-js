'use strict';

(function () {

    var hasLocalReference = /\=.*[a-zA-Z]+[0-9]+/;
    var hasForeignReference = /\=.*['].*[']![a-zA-Z]+[0-9]+/;
    var hasConstantFormula = /\=[0-9]?.*[a-zA-Z]+[0-9]?\(.*\)/;

    // Formats cell 
    function formatCell(cell, formula) {
        if (!formula) {
            // empty string
            return; 
        } 

        if (isNaN(formula)) {

            // If cell contains foreign reference 
            if (hasForeignReference.test(formula)) {
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
    }

    // Calls formatCell on all elements in the range 
    function processRange(range) {
        var formulas = range.formulas;
        var x = formulas.length;
        var y = formulas[0].length;

        for (var i = 0; i < x; i++) {
            for (var j = 0; j < y; j++) {
                formatCell(range.getCell(i,j), formulas[i][j]);
            }
        }
    }

    function colorize() {
        var t0 = performance.now();
        Excel.run(function (context) {
            var usedRange = context.workbook.getSelectedRange()
                .worksheet
                .getUsedRange();

            usedRange.load('formulas');

            return context.sync(usedRange)
                .then(processRange)
                .then(context.sync)
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

    //Assign function to DOM element
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#set-color').click(colorize);
        });
    };

})();
