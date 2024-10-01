/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import IVisualHost = powerbi.extensibility.IVisualHost;
import * as d3 from "d3";
type Selection<T extends d3.BaseType> = d3.Selection<T, any, any, any>;

export class Visual implements IVisual {
    private host: IVisualHost;
    private table: Selection<HTMLElement>;
    private tableHeader: Selection<HTMLElement>;
    private tableBody: Selection<HTMLElement>;

    constructor(options: VisualConstructorOptions) {
        this.table = d3.select(options.element)
            .append('table')
            .classed('simpleTable', true);
        
        this.tableHeader = this.table.append('thead')
            .append('tr');
        
        this.tableBody = this.table.append('tbody');
    }

    // Updating table
    
    public update(options: VisualUpdateOptions) {
        if (!options.dataViews || !options.dataViews[0]) return;
    
        let dataView: DataView = options.dataViews[0];
        let tableData = dataView.table;
    
        if (!tableData || !tableData.columns || !tableData.rows) {
            console.log("No data available for the table");
            return;
        }
    
        // Separate columns and rows based on the data roles
        let rowIndices = tableData.columns
            .map((col, i) => ({ index: i, role: col.roles }))
            .filter(col => col.role && col.role['rows']);
    
        let columnIndices = tableData.columns
            .map((col, i) => ({ index: i, role: col.roles }))
            .filter(col => col.role && col.role['columns']);
    
        let valueIndices = tableData.columns
            .map((col, i) => ({ index: i, role: col.roles }))
            .filter(col => col.role && col.role['values']);
    
        // Update header
        this.tableHeader.selectAll('th').remove();
        
        // Add header for row identifier
        this.tableHeader.append('th')
            .text(tableData.columns[rowIndices[0].index].displayName)
            .classed('rowHeader', true);

        // Add headers for columns and values
        this.tableHeader.selectAll('th.dataHeader')
            .data(columnIndices.concat(valueIndices))
            .enter()
            .append('th')
            .classed('dataHeader', true)
            .text(d => tableData.columns[d.index].displayName);
    
        // Update rows
        let rows = this.tableBody.selectAll('tr')
            .data(tableData.rows);
    
        rows.exit().remove();
    
        let newRows = rows.enter()
            .append('tr');
    
        let allRows = newRows.merge(rows as any);
    
        allRows.selectAll('td')
            .data(d => {
                return [
                    { value: d[rowIndices[0].index], isRowHeader: true },
                    ...columnIndices.concat(valueIndices).map(i => ({ value: d[i.index], isRowHeader: false }))
                ];
            })
            .join('td')
            .attr('class', d => d.isRowHeader ? 'rowHeader' : 'dataCell')
            .text(d => d.value !== null && d.value !== undefined ? d.value.toString() : "");
    
        // Apply styling
        this.table
            .style("border-collapse", "collapse")
            .style("width", "100%");
        
        this.table.selectAll("th, td")
            .style("border", "1px solid black")
            .style("padding", "5px")
            .style("text-align", "left");
        
        this.table.selectAll("th.rowHeader, td.rowHeader")
            .style("font-weight", "bold")
            .style("background-color", "#f2f2f2");

        // Log information for debugging
        console.log("Number of columns:", columnIndices.length + valueIndices.length + 1); // +1 for row header
        console.log("Number of rows:", tableData.rows.length);
    }
}