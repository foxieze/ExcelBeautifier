/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const formula = require("excel-formula")

Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "block";
        updateFormulaArea(null);
        Excel.run(function(context) {
            var workbook = context.workbook;
            workbook.onSelectionChanged.add(updateFormulaArea);

            return context.sync()
                .then(function() {
                    console.log("Event handler successfully registered for onChanged event in the worksheet.");
                });
        }).catch(error);
    }
});

function updateFormulaArea(event) {
    return Excel.run(function(context) {
        const range = context.workbook.getSelectedRange();
            // range.format.fill.color = "yellow";
            range.load("formulas");
            return context.sync().then(function() {
                try {
                    document.getElementById("area").innerHTML = formula.formatFormulaHTML(range.formulas.toString());
                } catch (error) {
                    console.log(error);
                }
            });
    })
} 

