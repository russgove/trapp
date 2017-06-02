import { TimeSpent } from '../dataModel';
import { ITrTimeCardState } from './ITrTimeCardState';
export interface ITrTimeCardProps {

    initialState: ITrTimeCardState,
    userName: string,
    userId: number, // id if user in the siteUsers list (a number)
    save: (ts: Array<TimeSpent>) => Promise<Array<TimeSpent>>,
    getTimeSpent: (weekEndingDate: Date) => Promise<Array<TimeSpent>>,
    editFormUrlFormat: string,
}
