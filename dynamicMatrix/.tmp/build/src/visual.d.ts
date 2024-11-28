import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private host;
    private table;
    private tableHeader;
    private tableBody;
    private valueFormats;
    constructor(options: VisualConstructorOptions);
    update(options: VisualUpdateOptions): void;
    private getCellValue;
    private getLeafNodes;
    private buildColumnHeaders;
    private calculateTotals;
}
