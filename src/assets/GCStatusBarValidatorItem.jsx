
import GC from "@grapecity/spread-sheets"

export default class GCStatusBarValidatorItem extends GC.Spread.Sheets.StatusBar.StatusItem {

    STATUSBAR_NS ='.statusbar';
    name= '';
    option= '';
    _context = GC.Spread.Sheets.Workbook || null;
    _eventNs=  '';
    _dispalyControl =  HTMLElement || null;
    _sheetMessage = "当前sheet共有错误数：";
    _bookMessage = " | 点击查看整个工作薄";
    constructor(name, options) {
        super(name, options);
        this.name = name;
        this.option = options;
        this._context = null;
        this._eventNs = this.STATUSBAR_NS + this.name;
        this._dispalyControl = null;
    }

    onCreateItemView(container) {
        let self = this;
        if (self && !self._dispalyControl) {
            self._dispalyControl = document.createElement('div');
            self._dispalyControl.style.color = "red";
            let sheetMessageSpan = document.createElement('span');
            let errorCountSpan = document.createElement('span');
            let booktMessageSpan = document.createElement('span');
            sheetMessageSpan.innerHTML = self._sheetMessage;
            errorCountSpan.innerHTML = "0";
            booktMessageSpan.innerHTML = self._bookMessage;
            self._dispalyControl.appendChild(sheetMessageSpan);
            self._dispalyControl.appendChild(errorCountSpan);
            self._dispalyControl.appendChild(booktMessageSpan);
            booktMessageSpan.onclick = function () {
                if (self.option.workbooMessageClick) {
                    self.option.workbooMessageClick.call(null, self._getWorkbookValidatorResult());
                }
            }
            container.appendChild(self._dispalyControl);
        }
    }

    onBind(context) {
        let self = this;
        this._context = context;
        // context.bind(GC.Spread.Sheets.Events.ActiveSheetChanged + self._eventNs, function () {
            // self._onActionChangeValue();
        // });
        context.bind(GC.Spread.Sheets.Events.RangeChanged + self._eventNs, function () {
            self._onActionChangeValue();
        });
        context.bind(GC.Spread.Sheets.Events.ValueChanged + self._eventNs, function () {
            self._onActionChangeValue();
        });
        // 监听fromJSON后事件 和 通过代码切换active sheet
        context.bind((GC.Spread.Sheets.Events).FormulatextboxActiveSheetChanged + self._eventNs, function (e, data) {
            // if (data.oldSheet === undefined) {
                self._onActionChangeValue();
            // }
        });


    }
    // onUpdate() {
    //     super.onUpdate();
    // }
    onUnbind() {
        if (this._context) {
            this._context.unbind(GC.Spread.Sheets.Events.ValueChanged + this._eventNs);
            this._context.unbind(GC.Spread.Sheets.Events.RangeChanged + this._eventNs);
            // this._context.unbind(GC.Spread.Sheets.Events.ActiveSheetChanged + this._eventNs);
            this._context.unbind((GC.Spread.Sheets.Events).FormulatextboxActiveSheetChanged + this._eventNs);
            this._context = null;
        }
    };
    onDispose() {
        this._context = null;
        if (this._dispalyControl) {
            this._dispalyControl.getElementsByTagName("span")[1].onclick = null;
        }
    }
    _onActionChangeValue() {
        if (this._dispalyControl != null && this._context) {
            let validatorResult = this._getSheetValidatorResult(this._context.getActiveSheet())
            this._dispalyControl.getElementsByTagName("span")[1].innerHTML = validatorResult.length.toString();
            console.log("_onActionChangeValue")
        }
    }

    _getSheetValidatorResult(sheet) {
        let rowCount = sheet.getRowCount(),
            colCount = sheet.getColumnCount();
        let validatorResult = [];
        for (let row = 0; row < rowCount; row++) {
            for (let col = 0; col < colCount; col++) {
                if (!sheet.isValid(row, col, sheet.getValue(row, col))) {
                    let dv = sheet.getDataValidator(row, col, GC.Spread.Sheets.SheetArea.viewport);
                    validatorResult.push(
                        {
                            sheetName: sheet.name(),
                            row: row,
                            col: col,
                            cell: GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1), 0, 0, GC.Spread.Sheets.CalcEngine.RangeReferenceRelative.allRelative, false),
                            title: dv.inputTitle(),
                            message: dv.inputMessage()
                        }
                    )
                }
            }
        }
        return validatorResult;
    }

    _getWorkbookValidatorResult() {
        let validatorResult = [];
        if (this._context) {
            for (let index = 0; index < this._context.getSheetCount(); index++) {
                let sheet = this._context.getSheet(index);
                let result = this._getSheetValidatorResult(sheet);
                if (result && result.length > 0) {
                    validatorResult = validatorResult.concat(result);
                }
            }
        }
        return validatorResult;
    }
}