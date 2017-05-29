import {TechnicalRequest} from '../dataModel';
import {ITrTimeCardState} from './ITrTimeCardState';
export interface ITrTimeCardProps {
   activeTRs:Array<TechnicalRequest>, // list of trs a person is working on
   initialState:ITrTimeCardState
}
