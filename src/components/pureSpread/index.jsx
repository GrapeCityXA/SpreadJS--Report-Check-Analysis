import React, { Component } from 'react'
import { SpreadSheets, Worksheet } from '@grapecity/spread-sheets-react';
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import { IO } from "@grapecity/spread-excelio"
import { Upload, Button, Select, message, Table, Tag, Space } from "antd"
import { UploadOutlined, DownloadOutlined, AccountBookOutlined } from '@ant-design/icons';
import FileSaver from 'file-saver';
import modelData from '../../assets/model';
import * as GC from "@grapecity/spread-sheets"
import PubSub from 'pubsub-js';
import "@grapecity/spread-sheets-print";
import configConst from '../../assets/\bcolConfig';
import dataSource from '../../assets/dataSource';
import { getKeyThenIncreaseKey } from 'antd/lib/message';
const { Option } = Select
export default class PureSpread extends Component {
    constructor(props) {
        super(props)
        this.spread = null

    }
    state = {
        fontFamily: null,
        configDatas: null,
        editSheet: null,
        modleName: null,
        tableData: [],
        openReportName: null,
    }
    // 初始化表格时
    spreadInitial = (spread) => {
        this.spread = spread

    }
    analysis = () => {
        debugger;
        window.location.href = window.location.origin + '/#/?model=' + this.state.openReportName;
    }

    importExcel = file => {

        let fileName = file.name.split(".")
        this.setState({
            modleName: fileName[0]
        })
        let excelio = new IO()
        excelio.open(file, (json) => {
            debugger;
            sessionStorage.setItem(this.state.modleName, JSON.stringify(json));
            this.setState({ openReportName: this.state.modleName });
            let tables = Object.keys(sessionStorage);
            let tableItem = []
            tables.forEach(item => {
                tableItem.push({
                    key: '1',
                    name: item,
                })
            })
            this.setState({
                tableData: tableItem
            })
            this.spread.fromJSON(json)
        })
        return false
    }
    // 精算
    calculate = (record) => {
        debugger;
        console.log(record);
        window.location.href = window.location.origin + '/#/?model=' + record.name;
    }
    handleExportExcel = () => {
        let excelJson = this.spread.toJSON()
        let excelio = new IO()
        excelio.save(excelJson, (blob) => {
            FileSaver.saveAs(blob, `${this.state.openReportName}.xlsx`)
        })
    }

    componentDidMount() {
        debugger;
        sessionStorage.setItem('财务报表', modelData)
        let tables = Object.keys(sessionStorage);
        this.spread.fromJSON(JSON.parse(modelData));
        this.setState({ openReportName: '财务报表' });
        let tableItem = []
        tables.forEach(item => {
            tableItem.push({
                key: '1',
                name: item,
            })
        })
        this.setState({
            tableData: tableItem
        })
    }

    render() {
        const props = {
            name: 'file',
            showUploadList: false,
            accept: ".xlsx",
            maxCount: 1,
            beforeUpload: this.importExcel
        }
        const columns = [
            {
                title: '精算模型名称',
                dataIndex: 'name',
                key: 'name',
                render: (text, record) => (<a
                    href="javascript:;"
                    onDoubleClick={() => this.calculate(record)}
                >{text}</a>),
            },
            {
                title: '操作',
                key: 'action',
                render: (text, record) => (
                    <Space size="middle">
                        <a
                            href="javascript:;"
                            onClick={() => this.calculate(record)}
                        >精算</a>
                    </Space>
                ),
            },
        ];

        const data = this.state.tableData
        return (
            <div style={{
                height: '100%'
            }}>
                <div className="buttons-div" style={{
                    height: '40px',
                    padding: '4px 0',
                }}>
                    <Upload {...props}>
                        <Button size="middle" icon={<UploadOutlined />}>导入财务报表</Button>
                    </Upload>
                    <Button icon={<DownloadOutlined />} size="middle" style={{ marginLeft: '12px' }} type="primary" onClick={this.handleExportExcel}>导出财务报表</Button>
                    <Button icon={<AccountBookOutlined />} type="primary" size="middle" style={{ marginLeft: '12px' }} onClick={this.analysis}>勾稽分析</Button>
                </div>
                <div style={{
                    height: 'calc(100% - 48px)',
                    // display: 'none'
                }}>
                    <SpreadSheets
                        workbookInitialized={this.spreadInitial}
                    ></SpreadSheets>
                </div>
                {/* <Table columns={columns} dataSource={data} /> */}
            </div>

        )
    }

}
