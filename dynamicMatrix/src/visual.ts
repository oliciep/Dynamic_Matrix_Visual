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
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import IVisualHost = powerbi.extensibility.IVisualHost;
import DataViewMatrix = powerbi.DataViewMatrix;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import * as d3 from "d3";
type Selection<T extends d3.BaseType> = d3.Selection<T, any, any, any>;

interface LeafNode {
    node: DataViewMatrixNode;
    levelValues: any[]; 
    index: number; 
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private table: Selection<HTMLTableElement>;
    private tableHeader: Selection<HTMLTableSectionElement>;
    private tableBody: Selection<HTMLTableSectionElement>;
    private valueFormats: string[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
    
        // Create a scrollable container
        let scrollContainer = d3.select(options.element)
            .append('div')
            .style('width', '100%')
            .style('max-height', '100%') // Set desired height
            .style('overflow', 'auto');
    
        // Append the table to the scrollable container
        this.table = scrollContainer.append('table')
            .classed('matrixTable', true);
    
        this.tableHeader = this.table.append('thead');
        this.tableBody = this.table.append('tbody');
    }

    public update(options: VisualUpdateOptions) {
        if (!options.dataViews || !options.dataViews[0]) return;
    
        let dataView: DataView = options.dataViews[0];
        let matrix: DataViewMatrix = dataView.matrix;
    
        if (!matrix || !matrix.rows || !matrix.columns) {
            console.log("No matrix data available");
            return;
        }
    
        // Collect format strings for value sources
        this.valueFormats = [];
        if (dataView.matrix.valueSources) {
            dataView.matrix.valueSources.forEach(valueSource => {
                this.valueFormats.push(valueSource.format || "");
            });
        }
    
        // Get leaf nodes with hierarchical values
        let rowLeaves = this.getLeafNodes(matrix.rows.root);
        let columnLeaves = this.getLeafNodes(matrix.columns.root);
    
        // Clear previous table content
        this.tableHeader.selectAll('*').remove();
        this.tableBody.selectAll('*').remove();
    
        // Build column headers
        let columnHeaders = this.buildColumnHeaders(columnLeaves);
    
        // Append column headers to the table
        columnHeaders.forEach(headerRowData => {
            let headerRow = this.tableHeader.append('tr');
            // Empty cells for row headers
            if (headerRowData.level === 0) {
                for (let i = 0; i < matrix.rows.levels.length; i++) {
                    headerRow.append('th')
                        .attr('rowspan', columnHeaders.length)
                        .classed('rowHeader', true);
                }
            }
            // Add column headers
            headerRow.selectAll('th.columnHeader')
                .data(headerRowData.items)
                .enter()
                .append('th')
                .attr('colspan', d => d.colspan)
                .text(d => d.text);
        });
    
        // Calculate totals
        let { columnTotals, grandTotal } = this.calculateTotals(rowLeaves, columnLeaves);
    
        // Build data rows
        rowLeaves.forEach((rowLeaf, rowIndex) => {
            let row = this.tableBody.append('tr');
    
            // Add row headers
            rowLeaf.levelValues.forEach(value => {
                row.append('td')
                    .classed('rowHeader', true)
                    .text(value != null ? value.toString() : "");
            });
    
            // Add data cells
            columnLeaves.forEach((columnLeaf, columnIndex) => {
                let cellValue = this.getCellValue(rowLeaf, columnLeaf);
                row.append('td')
                    .classed('dataCell', true)
                    .text(cellValue != null ? cellValue.toString() : "");
            });
        });
    
        // Add column totals row
        let totalsRow = this.tableBody.append('tr');
        totalsRow.append('td')
            .attr('colspan', matrix.rows.levels.length)
            .classed('totalsLabel', true)
            .text('Total');
    
        columnTotals.forEach(total => {
            totalsRow.append('td')
                .classed('columnTotal', true)
                .text(total.toString());
        });
    
        // Add grand total (DEPRECATED - WIP)
        // totalsRow.append('td')
        //    .classed('grandTotal', true)
        //    .text(grandTotal.toString());
    }

    // Helper method to get cell value
    private getCellValue(rowLeaf: LeafNode, columnLeaf: LeafNode): string | null {
        let rowNode = rowLeaf.node;
        let columnIndex = columnLeaf.index;
    
        if (rowNode.values && rowNode.values.hasOwnProperty(columnIndex)) {
            let valueObj = rowNode.values[columnIndex];
            if (valueObj && valueObj.value !== undefined) {
                // Get the measure index
                let measureIndex = valueObj.valueSourceIndex ?? 0; // Default to 0 if undefined
    
                // Get the format string for the measure
                let formatString = this.valueFormats[measureIndex] || "";
    
                // Create a value formatter
                let formatter = valueFormatter.create({
                    format: formatString,
                    value: valueObj.value
                });
    
                // Format and return the value
                return formatter.format(valueObj.value);
            }
        }
        return null;
    }

    // Recursive method to collect leaf nodes and their hierarchical values
    private getLeafNodes(node: DataViewMatrixNode, levelValues: any[] = [], leafNodes: LeafNode[] = [], level: number = 0): LeafNode[] {
        let currentLevelValues = [...levelValues];

        if (node.value !== undefined) {
            currentLevelValues.push(node.value);
        }

        if (node.children && node.children.length > 0) {
            node.children.forEach(child => {
                this.getLeafNodes(child, currentLevelValues, leafNodes, level + 1);
            });
        } else {
            leafNodes.push({
                node: node,
                levelValues: currentLevelValues,
                index: leafNodes.length // Assign index based on position
            });
        }

        return leafNodes;
    }
    

    // Build column headers with proper colspan
    private buildColumnHeaders(columnLeaves: LeafNode[]): { level: number, items: { text: string, colspan: number }[] }[] {
        // Build headers for each level
        let headersByLevel: { [level: number]: { [text: string]: { text: string, colspan: number } } } = {};

        columnLeaves.forEach(leaf => {
            leaf.levelValues.forEach((value, level) => {
                if (!headersByLevel[level]) {
                    headersByLevel[level] = {};
                }
                let key = value !== null && value !== undefined ? value.toString() : "";
                if (!headersByLevel[level][key]) {
                    headersByLevel[level][key] = { text: key, colspan: 0 };
                }
                headersByLevel[level][key].colspan += 1;
            });
        });

        // Convert headersByLevel to an array of header rows
        let headerRows: { level: number, items: { text: string, colspan: number }[] }[] = [];

        Object.keys(headersByLevel).forEach(levelKey => {
            let level = parseInt(levelKey);
            let items = Object.values(headersByLevel[level]);
            headerRows.push({ level: level, items: items });
        });

        // Sort header rows by level
        headerRows.sort((a, b) => a.level - b.level);

        return headerRows;
    }

        private calculateTotals(rowLeaves: LeafNode[], columnLeaves: LeafNode[]): { columnTotals: number[], grandTotal: number } {
        let columnTotals = new Array(columnLeaves.length).fill(0);
        let grandTotal = 0;
    
        rowLeaves.forEach((rowLeaf, rowIndex) => {
            columnLeaves.forEach((columnLeaf, columnIndex) => {
                let cellValue = this.getCellValue(rowLeaf, columnLeaf);
                let numericValue = parseFloat(cellValue) || 0;
                columnTotals[columnIndex] += numericValue;
                grandTotal += numericValue;
            });
        });
    
        return { columnTotals, grandTotal };
    }
}
