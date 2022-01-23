import React, { Component,Suspense} from 'react'
import { Switch,Route,HashRouter, Redirect} from 'react-router-dom';
import '../../App.less'
import PureSpread from'../pureSpread';
import SpreadDesigner from '../designer'
// const PureSpread = React.lazy(() => import('../pureSpread'))
// const SpreadDesigner = React.lazy(() => import('../designer'))
export default class MiddleComponent extends Component {
    state = {
        configDatas: null,
    }
    getDatas=(msg)=>{
        this.setState({
            configDatas:msg
        });
    }
    render() {
        return (
            <div className="index-content" style = {{height:'95%'}}>
                <HashRouter>
                    <Switch>
                        <Suspense fallback={<div>loading...</div>}>
                            <Route exact path="/designer" component={PureSpread}><PureSpread configDatas = {this.state.configDatas}/></Route>
                            <Route exact path="/" component={SpreadDesigner}><SpreadDesigner getDatas={this.getDatas.bind(this)}/></Route>
                        </Suspense>
                    </Switch>
                </HashRouter>
            </div>
        )
    }
}
