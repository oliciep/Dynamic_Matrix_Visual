import { Visual } from "../../src/visual";
import powerbiVisualsApi from "powerbi-visuals-api";
import IVisualPlugin = powerbiVisualsApi.visuals.plugins.IVisualPlugin;
import VisualConstructorOptions = powerbiVisualsApi.extensibility.visual.VisualConstructorOptions;
import DialogConstructorOptions = powerbiVisualsApi.extensibility.visual.DialogConstructorOptions;
var powerbiKey: any = "powerbi";
var powerbi: any = window[powerbiKey];
var dynamicMatrixFDD1BDC1563544A6A39751B0CF710976_DEBUG: IVisualPlugin = {
    name: 'dynamicMatrixFDD1BDC1563544A6A39751B0CF710976_DEBUG',
    displayName: 'DynamicMatrix',
    class: 'Visual',
    apiVersion: '5.3.0',
    create: (options?: VisualConstructorOptions) => {
        if (Visual) {
            return new Visual(options);
        }
        throw 'Visual instance not found';
    },
    createModalDialog: (dialogId: string, options: DialogConstructorOptions, initialState: object) => {
        const dialogRegistry = (<any>globalThis).dialogRegistry;
        if (dialogId in dialogRegistry) {
            new dialogRegistry[dialogId](options, initialState);
        }
    },
    custom: true
};
if (typeof powerbi !== "undefined") {
    powerbi.visuals = powerbi.visuals || {};
    powerbi.visuals.plugins = powerbi.visuals.plugins || {};
    powerbi.visuals.plugins["dynamicMatrixFDD1BDC1563544A6A39751B0CF710976_DEBUG"] = dynamicMatrixFDD1BDC1563544A6A39751B0CF710976_DEBUG;
}
export default dynamicMatrixFDD1BDC1563544A6A39751B0CF710976_DEBUG;