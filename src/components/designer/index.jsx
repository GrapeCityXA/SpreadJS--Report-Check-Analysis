import React, { Component } from 'react'
import '@grapecity/spread-sheets-designer-resources-cn';
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css"
import '@grapecity/spread-sheets-designer/styles/gc.spread.sheets.designer.min.css'
import { Designer } from '@grapecity/spread-sheets-designer-react';
import * as designerGC from '@grapecity/spread-sheets-designer';
import GC from '@grapecity/spread-sheets';
import * as echarts from 'echarts';
import {
    Button, Switch, Drawer, Form, Input, Checkbox, Row, Col,
    Select, Radio, InputNumber, Spin, Popconfirm, message
} from "antd"
import { NodeCollapseOutlined, CheckCircleOutlined, NodeExpandOutlined, ExpandAltOutlined } from '@ant-design/icons';
import './designer.less';
import GCStatusBarValidatorItem from '../../assets/GCStatusBarValidatorItem'
const { Option } = Select;

export default class SpreadDesigner extends Component {
    constructor(props) {
        super(props)
        this.designer = null
    }
    state = {
        config: designerGC.Spread.Sheets.Designer.DefaultConfig,
        colInfos: [],
        showFormula: true,
        drawerVisible: false,
        loading: false,
        trackType: '',
        trackCellInfo: '',
        rootNode1: '',
        precedentsRootNode: '',
        dependentsRootNode: ''
    }
    // 初始化编辑器
    designerInitial = (designer) => {
        this.designer = designer;
        this.designer.setConfig(GC.Spread.Sheets.Designer.ToolBarModeConfig);
        this.designer.getWorkbook().getActiveSheet().options.showFormulas = true;
        // this.designer.getWorkbook().getActiveSheet().suspendCalcService(false);
        this.designer.getWorkbook().bind(GC.Spread.Sheets.Events.SheetTabClick, (e, info) => {
            // 显示公式
            // info.sheet.options.showFormulas = this.state.showFormula;
            // 挂起计算服务
            // info.sheet.suspendCalcService(false)
        })
    }
    trackPrecedentsCell = () => {
        let spread = this.designer.getWorkbook();
        let sheet = this.designer.getWorkbook().getActiveSheet();
        let trackType = "Precedents";
        let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
        this.trackCellInfoChanged(trackCellInfo, spread, trackType)
        // this.setState({rootNode1:'55'});
    }
    trackDependentsCell = () => {
        let spread = this.designer.getWorkbook();
        let sheet = this.designer.getWorkbook().getActiveSheet();
        let trackType = "Dependents";
        let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
        this.trackCellInfoChanged(trackCellInfo, spread, trackType)
    }
    trackAllCell = () => {
        let spread = this.designer.getWorkbook();
        let sheet = this.designer.getWorkbook().getActiveSheet();
        let trackType = "Both";
        let trackCellInfo = sheet.name() + "*" + sheet.getActiveRowIndex() + "*" + sheet.getActiveColumnIndex() + "*" + Math.random();
        this.trackCellInfoChanged(trackCellInfo, spread, trackType)
    }
    reviewAllCells = () => {
        this.state.trackType = 'review'
        let statusBar = GC.Spread.Sheets.StatusBar.findControl(document.getElementsByClassName("gc-statusBar"));
        // 添加数据验证状态栏
        statusBar.add(new GCStatusBarValidatorItem("validatorStatus", { tipText: "数据验证状态", workbooMessageClick: this.getWorkbookValidatorResult }), 1);
        this.getWorkbookValidatorResult();
    }
    trackCellInfoChanged = (trackCellInfo, sourceSpread, trackType) => {
        this.state.trackType = trackType;
        this.state.trackCellInfo = trackCellInfo;
        this.sourceSpread = sourceSpread;
        if (trackCellInfo && sourceSpread) {
            this.buildNodeTreeAndPaint(sourceSpread, trackCellInfo);
        }
    }
    // 递归构建追踪树
    buildNodeTreeAndPaint = (spreadSource, trackCellInfo) => {
        let info = this.getCellInfo(trackCellInfo);
        let sheetSource = spreadSource.getSheetFromName(info.sheetName);
        // 创建跟节点
        let rootNode = this.creatNode(info.row, info.col, sheetSource, 0, "");

        let name = rootNode.sheetName + "*" + rootNode.row + "*" + rootNode.col + "*" + Math.random().toString();
        let precedentsRootNode = '';
        let dependentsRootNode = '';
        if (this.state.trackType === "Precedents" || this.state.trackType === "Both") {
            this.getNodeChild(rootNode, sheetSource, "Precedents")
            debugger;
            console.log(rootNode)
            if (this.state.trackType === "Both") {
                let rootNodeChildren = JSON.parse(JSON.stringify(rootNode.children));
                rootNode.children = [];
                precedentsRootNode = JSON.parse(JSON.stringify(rootNode));
                precedentsRootNode.children.push({
                    name: "Precedents",
                    value: "Precedents",
                    children: rootNodeChildren
                })
                this.setState({
                    precedentsRootNode: JSON.parse(JSON.stringify(precedentsRootNode)),
                })
            }
        }
        if (this.state.trackType === "Dependents" || this.state.trackType === "Both") {
            this.getNodeChild(rootNode, sheetSource, "Dependents")
            console.log(rootNode)
            if (this.state.trackType === "Both") {
                let deepInfo = [1];
                let rootNodeChildren = JSON.parse(JSON.stringify(rootNode.children));
                rootNode.children = [];
                dependentsRootNode = JSON.parse(JSON.stringify(rootNode));
                dependentsRootNode.children.push({
                    name: "Dependents",
                    value: "Dependents",
                    children: rootNodeChildren
                })
                this.setState({
                    dependentsRootNode: JSON.parse(JSON.stringify(dependentsRootNode)),
                })
            }



        }
        if (this.state.trackType === "Both") {
            precedentsRootNode.children = precedentsRootNode.children.concat(dependentsRootNode.children);
            // let bothRootNode = precedentsRootNode.children[0].children.concat(dependentsRootNode.children[0].children)
            this.setState({
                rootNode1: JSON.parse(JSON.stringify(precedentsRootNode)),
            })
        } else {
            this.setState({
                rootNode1: JSON.parse(JSON.stringify(rootNode)),
            })
        }
    }
    creatNode = (row, col, sheet, deep, trackType) => {
        let node = {
            value: sheet.getValue(row, col),
            position: sheet.name() + "!" + GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1)),
            deep: deep,
            name: `${sheet.name()}!${GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1))}\nvalue:${sheet.getValue(row, col)}`,
            sheetName: sheet.name(),
            row: row,
            col: col,
            trackType: trackType
        };
        return node;
    }
    getNodeChild = (rootNode, sheet, trackType) => {
        let childNodeArray = [];
        let children = [];
        let row = rootNode.row, col = rootNode.col, deep = rootNode.deep;
        if (trackType == "Precedents") {
            children = sheet.getPrecedents(row, col);
        }
        else {
            children = sheet.getDependents(row, col);
        }
        // let self = this;
        if (children.length >= 1) {
            children.forEach((node) => {
                let row = node.row,
                    col = node.col,
                    rowCount = node.rowCount,
                    colCount = node.colCount,
                    _sheet = sheet.parent.getSheetFromName(node.sheetName);
                if (rowCount > 1 || colCount > 1) {
                    for (let r = row; r < row + rowCount; r++) {
                        for (let c = col; c < col + colCount; c++) {
                            let newNode = this.creatNode(r, c, _sheet, deep + 1, trackType)
                            // if (deep < self.maxDeep) {
                            this.getNodeChild(newNode, _sheet, trackType);
                            // }
                            childNodeArray.push(newNode);
                        }
                    }
                } else {
                    let newNode = this.creatNode(row, col, _sheet, deep + 1, trackType)
                    // if (deep < self.maxDeep) {
                    this.getNodeChild(newNode, _sheet, trackType);
                    // }
                    childNodeArray.push(newNode);
                }
            });
        }
        rootNode.children = childNodeArray;
    }
    
    getCellInfo = (cellInfo) => {
        let info = cellInfo.split("*");
        return {
            sheetName: info[0],
            row: parseInt(info[1]),
            col: parseInt(info[2])
        }
    }
    showNodePositon = (param) => {

        this.designer.getWorkbook().setActiveSheet(param.data.sheetName);
        let sheet = this.designer.getWorkbook().getActiveSheet();
        if (param.data.row) {
            sheet.setActiveCell(param.data.row, param.data.col);
            sheet.showCell(param.data.row, param.data.col, GC.Spread.Sheets.VerticalPosition.center, GC.Spread.Sheets.HorizontalPosition.center);
            console.log(param);
        }


    }
    getSheetValidatorResult = (sheet)=> {
        let rowCount = sheet.getRowCount(),
            colCount = sheet.getColumnCount();
        let validatorResult = [];
        let sheetErrorData = {
            name:sheet.name(),
            children:[],
            itemStyle:{
                color:'red'
            },
            lineStyle:{
                color:'red'
            },
        }
        for (let row = 0; row < rowCount; row++) {
            for (let col = 0; col < colCount; col++) {
                if (!sheet.isValid(row, col, sheet.getValue(row, col))) {
                    let dv = sheet.getDataValidator(row, col, GC.Spread.Sheets.SheetArea.viewport);
                    validatorResult.push(
                        {
                            name:`${sheet.name()}!${GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1))}\nvalue:${sheet.getValue(row, col)}`,
                            value:sheet.getValue(row, col),
                            sheetName: sheet.name(),
                            row: row,
                            col: col,
                            cell: GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1), 0, 0, GC.Spread.Sheets.CalcEngine.RangeReferenceRelative.allRelative, false),
                            title: dv.inputTitle(),
                            message: dv.inputMessage(),
                            itemStyle:{
                                color:'red'
                            },
                            lineStyle:{
                                color:'red'
                            },
                        }
                    )
                }
            }
        }
        sheetErrorData.children = validatorResult;
        

        return sheetErrorData;
    }

    getWorkbookValidatorResult=()=> {
        let workbookErrorData = {
            name:'当前工作簿所有错误',
            children:[],
            itemStyle:{
                color:'red'
            },
            lineStyle:{
                color:'red'
            },
        }
        debugger;
        let validatorResult = [];
        if (this.designer.getWorkbook()) {
            for (let index = 0; index < this.designer.getWorkbook().getSheetCount(); index++) {
                let sheet = this.designer.getWorkbook().getSheet(index);
                let result = this.getSheetValidatorResult(sheet);
                if (result && result.children.length > 0) {
                    validatorResult = validatorResult.concat(result);
                }
            }
        }
        debugger;
        workbookErrorData.children = validatorResult
        this.setState({
            rootNode1: JSON.parse(JSON.stringify(workbookErrorData)),
        })
        
    }
    calculate = () => {
        debugger;

        console.log(this.designer.getWorkbook().getDataSource())
        // this.designer.getWorkbook().getActiveSheet().resumeCalcService(true);
    }

    showFormula = (checked) => {
        debugger;
        this.setState({
            showFormula: checked
        })
        this.designer.getWorkbook().getActiveSheet().options.showFormulas = checked;

    }

    onAgeChange = () => {

    }
    onTimeChage = () => {

    }
    saveCalculatorSpread = () => {
        debugger;
        let newData = JSON.stringify(this.designer.getWorkbook().toJSON())
        document.getElementsByClassName("gc-ribbon-bar")[0].classList.add("collapsed");
        let ref = window.location.href;
        if (ref.indexOf("?") !== -1) {
            let modelName = ref.split("?")[1].split("=")[1];
            sessionStorage.setItem(decodeURI(modelName), newData)
            alert("保存成功")
        }
    }
    componentDidUpdate() {
        debugger;
        document.getElementsByClassName("gc-ribbon-bar")[0].classList.add("collapsed")
        let chartDom = document.getElementById('main');
        let myChart = echarts.init(chartDom);
        let ROOT_PATH =
            'https://cdn.jsdelivr.net/gh/apache/echarts-website@asf-site/examples';


        let option;
        myChart.showLoading();
        myChart.hideLoading();
        myChart.setOption(
            (option = {
                tooltip: {
                    trigger: 'item',
                    triggerOn: 'mousemove'
                },
                series: [
                    {
                        type: 'tree',
                        data: [this.state.rootNode1],
                        top: '1%',
                        left: '15%',
                        bottom: '1%',
                        right: '7%',
                        symbolSize: 10,
                        orient: this.state.trackType === 'review'?'LR':'RL',
                        label: {
                            position: this.state.trackType === 'review'?'left':'right',
                            verticalAlign: 'middle',
                            align: this.state.trackType === 'review'?'right':'left',
                        },
                        leaves: {
                            label: {
                                position: this.state.trackType === 'review'?'right':'left',
                                verticalAlign: 'middle',
                                align: this.state.trackType === 'review'?'left':'right'
                            }
                        },
                        emphasis: {
                            focus: 'descendant'
                        },
                        // layout: 'radial',
                        expandAndCollapse: true,
                        animationDuration: 550,
                        animationDurationUpdate: 750
                    }
                ]
            })
        );

        option && myChart.setOption(option);

    }
    componentWillReceiveProps() {
        debugger;
    }
    componentDidMount() {
        debugger;
        let chartDom = document.getElementById('main');
        let myChart = echarts.init(chartDom);
        let ROOT_PATH =
            'https://cdn.jsdelivr.net/gh/apache/echarts-website@asf-site/examples';


        let option;
        myChart.showLoading();
        myChart.hideLoading();
        myChart.setOption(
            (option = {
                tooltip: {
                    trigger: 'item',
                    triggerOn: 'mousemove'
                },
                series: [
                    {
                        type: 'tree',
                        data: [this.state.rootNode1],
                        top: '1%',
                        left: '15%',
                        bottom: '1%',
                        right: '7%',
                        symbolSize: 7,
                        orient: 'RL',
                        label: {
                            position: 'right',
                            verticalAlign: 'middle',
                            align: 'left'
                        },
                        leaves: {
                            label: {
                                position: 'left',
                                verticalAlign: 'middle',
                                align: 'right'
                            }
                        },
                        emphasis: {
                            focus: 'descendant'
                        },
                        expandAndCollapse: false,
                        initialTreeDepth: 15,
                        animationDuration: 550,
                        animationDurationUpdate: 750
                    }
                ]
            })
        );

        option && myChart.setOption(option);
        myChart.on('click', this.showNodePositon)
        debugger;
        document.getElementsByClassName("gc-ribbon-bar")[0].classList.add("collapsed");
        let ref = window.location.href;
        if (ref.indexOf("?") !== -1) {
            let modelName = ref.split("?")[1].split("=")[1];
            let tableJson = sessionStorage.getItem(decodeURI(modelName));
            let spread = this.designer.getWorkbook();
            spread.fromJSON(JSON.parse(tableJson))
        }

        this.designer.getWorkbook().getActiveSheet().options.showFormulas = true;
    }
    // 组件卸载时上传配置表格，向数据展示组件传值
    componentWillUnmount() {

    }
    render() {

        const { config } = this.state
        return (
            <div style={{ height: '100%' }}>
                <Button icon={<NodeCollapseOutlined />} size="middle" style={{ margin: "10px 10px" }} type="primary" onClick={this.trackPrecedentsCell} >追踪引用单元格</Button>
                <Button icon={<NodeExpandOutlined />} size="middle" style={{ margin: "10px 10px" }} type="primary" onClick={this.trackDependentsCell} >追踪从属单元格</Button>
                <Button icon={<ExpandAltOutlined />} size="middle" style={{ margin: "10px 10px" }} type="primary" onClick={this.trackAllCell} >追踪所有单元格</Button>
                <Popconfirm
                    title="请在工作簿状态栏查看工作表审查结果，所有审查结果请看右侧窗口！"
                    onConfirm={this.reviewAllCells}
                    placement="rightTop"
                    cancelText = '关闭'
                    okText="知道了"
                >
                    <Button icon={<CheckCircleOutlined />} size="middle" style={{ margin: "10px 10px" }} type="primary"  >审查所有单元格</Button>
                </Popconfirm>
                <span>显示公式：</span><Switch defaultChecked onChange={this.showFormula} />
                <Spin spinning={this.state.loading} delay={500} tip="计算中...">
                </Spin>
                <Row gutter={24} style={{ height: '99%', flexFlow: 'row !important' }}>
                    <Col span={12} key={1} style={{ paddingRight: '2px', height: '99%' }}>
                        <div style={{ border: '2px solid #eef0f3', height: '100%' }}>
                            <Designer
                                styleInfo={{
                                    height: '100%',
                                }}
                                designerInitialized={this.designerInitial}
                                config={config}
                            />
                        </div>
                    </Col>
                    <Col span={12} key={2} style={{ paddingLeft: '2px', height: '99%' }}>
                        <div id="main" style={{ height: '100%', border: '2px solid #eef0f3', }}>
                        </div>
                    </Col>
                </Row>
            </div>
        )
    }
}
