import { ITrTimeCardWebPartProps } from './ITrTimeCardWebPartProps';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { TimeSpent, TechnicalRequest } from "./dataModel";
import { Version } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as moment from "moment";
import * as strings from 'trTimeCardStrings';
import TrTimeCard from './components/TrTimeCard';
import { ITrTimeCardProps } from './components/ITrTimeCardProps';


import * as _ from 'lodash';
export default class TrTimeCardWebPart extends BaseClientSideWebPart<ITrTimeCardWebPartProps> {

  private reactElement: React.ReactElement<ITrTimeCardProps>;
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });
    });
  }

  /**
   * Adds a new row to the TimeSpent list
   * 
   * @param {any} batch The odata batch to execute the call in (from pnp.SP.createBatch)
   * @param {TimeSpent} timeSpent  A TimeSpent reord to be added to the TImeSpent list
   * @returns {Promise<number>}  The UID of the TimeSpent record that was edded
   * 
   * @memberof TrTimeCardWebPart
   */
  public AddTimeSpent(batch, timeSpent: TimeSpent): Promise<number> {
    return pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.inBatch(batch).add({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {
   
      return response.data.Id;
      // CAPTURE ID TO RETURN
    });

  }

  /**
   * Updates an existing record in the TimeSpent list
   * 
   * @param {any} batch  batch The odata batch to execute the call in (from pnp.SP.createBatch)
   * @param {TimeSpent} timeSpent  A TimeSpent reord to be added to the TImeSpent list. The Id must be present on this item
   * 
   * @memberof TrTimeCardWebPart
   */
  public UpdateTimeSpent(batch, timeSpent: TimeSpent) {
    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.getById(timeSpent.tsId).inBatch(batch).update({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {

    })
      .catch((error) => {
        console.log("ERROR, An error occured Updateing TimeSPent");
        console.log(JSON.stringify(error));
      });
  }

  /**
   *  Called by the TRTimeCard react Component, this method saves changes to the TimeSPent list. Records with no id are 
   * added. Record with an id are updated.
   * 
   * @param {Array<TimeSpent>} timeSpents -- the arry of timespents to be saved
   * @returns {Promise<Array<TimeSpent>>} -- returns an updated array of TimeSpents. record which wer added to the list
   * will have their ID's added.  TRTimeCard react Component sets its state to this new array after the update is made.
   * 
   * @memberof TrTimeCardWebPart
   */
  public save(timeSpents: Array<TimeSpent>): Promise<Array<TimeSpent>> {
    let batch = pnp.sp.createBatch();
    for (const timeSpent of timeSpents) {
      if (timeSpent.tsId === null) {
        if (timeSpent.hoursSpent !== 0) {
          this.AddTimeSpent(batch, timeSpent).then((id) => {
            timeSpent.tsId = id;
          });
        }
      }
      else {
        this.UpdateTimeSpent(batch, timeSpent);
      }

    }
    return batch.execute().then((x) => {
      return timeSpents;
    });
  }

  /**
   * Gets the TRs assigned to the current user that are not Completed..
   * The initial Timesheet grid displays all trs assigned to the user that are not completed (this list) as well as any 
   * TimeSpent record entered for the selectd week-ending date.
   * Note that when I get the TimeSpent records in getExistingTimeSpent, I cannot get the 'Ststus' field because we cannot
   * expand lookup columns with odata. So We get all the tomespent record for the selected week ending date, and then in 
   * a secon batch (trBatch) we go back an get the TRs associated with the existing TimeSpent records, and then update the 
   * TimeSpent records with the info from the TR.
   * 
   * @param {*} [batch]  batch The odata batch to execute the call in (from pnp.SP.createBatch)
   * @returns {Promise<Array<TechnicalRequest>>}  The Technical requests (TRs) assigned to the current user which are not completed.
   * 
   * @memberof TrTimeCardWebPart
   */
  public getAssignedTrs(batch?: any): Promise<Array<TechnicalRequest>> {

    // get the Active TRS Assigned to the user. These need to be shown in the timesheet
    let filterString = `(TRAssignedTo/EMail eq '${this.context.pageContext.user.email}') and (TRStatus ne 'Completed') and (TRStatus ne 'Canceled')`;
    let command = pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.expand("TRAssignedTo")
      .select("Title,RequestTitle,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
      .filter(filterString)
      .orderBy('RequiredDate');
    if (batch) {
      command.inBatch(batch);
    }
    return command.get().then((items) => {
      return _.map(items, (item) => {
        return {
          trId: item["Id"],
          title: item["Title"],
          status: item["TRStatus"],
          requiredDate: item["RequiredDate"],
          priority: item["TRPriority"],
          requestTitle: item["RequestTitle"],
        };
      });
    });
  }

  /**
   * Gets the speficied tr from the TR list
   * We call this for each TimeSpent record to get the additional TR metadata for a Timespent record.
   * @param {number} trId  the ID of the tr to fetch
   * @param {*} [batch]  The odata batch to execute the call in (from pnp.SP.createBatch)
   * @returns {Promise<TechnicalRequest>}  The techinal requyest record(only partially filled in with the neccesary metadata to show
   * in the Timesheet grid)
   * 
   * @memberof TrTimeCardWebPart
   */
  public getTR(trId: number, batch?: any): Promise<TechnicalRequest> {
    let command = pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items
      .getById(trId)
      .expand("TRAssignedTo")
      .select("Title,RequestTitle,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail");

    if (batch) {
      command.inBatch(batch);
    }
    return command.get().then((item) => {

      return {
        trId: item["Id"],
        title: item["Title"],
        requestTitle: item["RequestTitle"],
        status: item["TRStatus"],
        requiredDate: item["RequiredDate"],
        priority: item["TRPriority"],
      };
    });
  }

  /**
   * Gets the TimeSpent record for the current user in the selecting week ending data. Note that each site must have
   * it's timezone set to GMT for the week-ending logic to work. We use moment to get the week ending date.
   * 
   * @param {Date} weekEndingDate The week ending date to fetch data for
   * @param {*} [batch] 
   * @returns {Promise<Array<TimeSpent>>} 
   * 
   * @memberof TrTimeCardWebPart
   */
  public getExistingTimeSpent(weekEndingDate: Date, batch?: any): Promise<Array<TimeSpent>> {
    let tsFilter = `(TechmicalSpecialist/EMail eq '${this.context.pageContext.user.email}') and (WeekEndingDate eq datetime'${weekEndingDate.toISOString()}')`;
    let command = pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.expand("TechmicalSpecialist,TR")
      .select("Id,WeekEndingDate,TRId,HoursSpent,TechmicalSpecialist/Id,TechmicalSpecialist/EMail,TR/Title,TR/Id,TR/RequiredDate")
      .filter(tsFilter);
    if (batch) {
      command.inBatch(batch);
    }
    return command.get()
      .then((items) => {
        return _.map(items, (item) => {

          return {
            tsId: item["Id"],
            technicalSpecialist: item["TechnicalSpecialist"],
            trId: item["TRId"],
            weekEndingDate: item["WeekEndingDate"],
            hoursSpent: item["HoursSpent"],
            trTitle: null,
            trRequestTitle: null,
            trStatus: null,
            trPriority: null,
            trRequiredDate: null,
          };
        });
      });

  }

  /**
   * This is the main method used to fetch data to be displayed in the grid. It calls getAssignedTrs, getExistingTimeSpent,
   * and getTR, and merges all the info together into an array of TimeSpent that gets passed to the react component 
   * to display. This method is called from the Render method of the webpart to get the initial data to display, 
   * and als when the user changes the date in the UI.
   * 
   * @param {Date} date -- the date to fetch timeSpent record s for
   * @returns {Promise<Array<TimeSpent>>}  An aArray of TimeSpent records to be displayed (includes infr from the TR as well)
   * 
   * @memberof TrTimeCardWebPart
   */
  public getTimeSpent(date: Date): Promise<Array<TimeSpent>> {
    // mainBatch is used to fetch TimeSpents and users TRs
    let mainBatch = pnp.sp.createBatch();
    let activeTRs: Array<TechnicalRequest> = [];
    let timeSpents: Array<TimeSpent> = [];
    let userId: number;
    pnp.sp.web.currentUser.inBatch(mainBatch).get()
      .then((user) => {
        userId = user.Id;
      });
    // get the Active TRS Assigned to the user. These need to be shown in the timesheet
    this.getAssignedTrs(mainBatch).then((items) => {
      activeTRs = items;
    });
    // get the Existing TimeSpents for the user in the selected weeek
    this.getExistingTimeSpent(date, mainBatch).then(items => {
      timeSpents = items;
    });
    return mainBatch.execute()
      .then((data) => {
        // trBatch is used to fetch TRS associated with the timeSPents, needs to execute after we get all  the timespents
        let trBatch = pnp.sp.createBatch();
        // 
        for (let timeSpent of timeSpents) {
          this.getTR(timeSpent.trId, trBatch)
            .then((tr) => {
              timeSpent.trPriority = tr.priority;
              timeSpent.trStatus = tr.status;
              timeSpent.trTitle = tr.title;
              timeSpent.trRequestTitle = tr.requestTitle;
              timeSpent.trRequiredDate = tr.requiredDate;
            })
            .catch((error) => {
              console.log("ERROR, An error occured fetching the TRs for the TS");
              console.log(error.message);
            });
        }
        // add any trs user has not reported time for yet
        return trBatch.execute().then((x) => {
          for (const tr of activeTRs) {
            // add a row for any active projects not on list
            const itemIndex = _.findIndex(timeSpents, (item) => { return item.trId === tr.trId; });
            if (itemIndex === -1) {
              timeSpents.push({
                trId: tr.trId,
                technicalSpecialist: userId,
                weekEndingDate: date,
                hoursSpent: 0,
                tsId: null,
                trTitle: tr.title,
                trRequestTitle: tr.requestTitle,
                trStatus: tr.status,
                trRequiredDate: tr.requiredDate,
                trPriority: tr.priority
              });
            }
            else {
              timeSpents[itemIndex].trPriority = tr.priority;
              timeSpents[itemIndex].trRequiredDate = tr.requiredDate;
              timeSpents[itemIndex].trStatus = tr.status;

            }
          }
          return timeSpents;
        });
      });
  }

  /**
   * SPFX renbder method 
   * gets data , Sets props and renders component
   * 
   * @memberof TrTimeCardWebPart
   */
  public render(): void {
    var defaultWeekEndDate: Date = new Date(moment().utc().endOf('isoWeek').startOf('day'));
    let props: ITrTimeCardProps = {
      userName: this.context.pageContext.user.displayName,
      userId: null,
      getTimeSpent: this.getTimeSpent.bind(this),
      save: this.save.bind(this),
      editFormUrlFormat: this.properties.editFormUrlFormat,
      webUrl: this.context.pageContext.web.absoluteUrl,
      initialState: {
        weekEndingDate: defaultWeekEndDate,
        timeSpents: [],
        message: "",
      }
    };
    // 
    this.getTimeSpent(defaultWeekEndDate).then((ts: Array<TimeSpent>) => {
      props.initialState.timeSpents = ts;
      this.reactElement = React.createElement(TrTimeCard, props);
      var formComponent: TrTimeCard = ReactDom.render(this.reactElement, this.domElement) as TrTimeCard;//render the component
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('editFormUrlFormat', {
                  label: "Url format for edit form"
                }),
                PropertyPaneTextField('technicalRequestListName', {
                  label: "List Name for Technical Requests"
                }),
                PropertyPaneTextField('timeSpentListName', {
                  label: "List Name for Time Spent"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
