import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from "sp-pnp-js";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { TR, WorkType, ApplicationType, EndUse } from "./dataModel";
import * as strings from 'trFormStrings';
import * as _ from 'lodash';
import TrForm from './components/TrForm';
import { ITrFormProps } from './components/ITrFormProps';
import { ITrFormWebPartProps } from './ITrFormWebPartProps';

export default class TrFormWebPart extends BaseClientSideWebPart<ITrFormWebPartProps> {
  private tr: TR;
  private workTypes: Array<WorkType> = [];
  private applicationTypes: Array<ApplicationType> = [];
  private endUses: Array<EndUse> = [];
  private reactElement: React.ReactElement<ITrFormProps>;
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });
      //  this.loadData();
    });
  }
  public render(): void {

    let formProps: ITrFormProps = { workTypes: [], applicationTypes: [], endUses: [], tr: new TR(), save: this.save };
    let batch = pnp.sp.createBatch();
    pnp.sp.web.lists.getByTitle("End Uses").items.inBatch(batch).get()
      .then((items) => {
        formProps.endUses = _.map(items, (item) => {
          return new EndUse(item["Id"], item["Title"], item["ApplicationTypeId"]);
        });
      });
    pnp.sp.web.lists.getByTitle("Work Types").items.inBatch(batch).get()
      .then((items) => {
        formProps.workTypes = _.map(items, (item) => {
          return new WorkType(item["Id"], item["Title"]);
        });

      });
    pnp.sp.web.lists.getByTitle("Application Types").items.inBatch(batch).get()
      .then((items) => {

        formProps.applicationTypes = _.map(items, (item) => {
          return new ApplicationType(item["Id"], item["Title"], item["WorkTypesId"]);
        });

      });
    // how to get querystring parameter
    pnp.sp.web.lists.getByTitle("trs").items.getById(1).inBatch(batch).get()
      .then((item) => {

        formProps.tr = new TR();
        formProps.tr.Id = item.Id;
        formProps.tr.ActualCompletionDate = item.ActualCompletionDate;
        formProps.tr.ApplicationTypeId = item.ApplicationTypeId;
        formProps.tr.ActualStartDate = item.ActualStartDate;
        formProps.tr.CER = item.CER;
        formProps.tr.Customer = item.Customer;
        formProps.tr.TRDueDate = item.TRDueDate;
        formProps.tr.EstimatedHours = item.EstimatedHours;
        formProps.tr.InitiationDate = item.InitiationDate;
        formProps.tr.MailBox = item.MailBox;
        formProps.tr.TRPriority = item.TRPriority;
        formProps.tr.Requestor = item.Requestor;
        formProps.tr.Site = item.Site;
        formProps.tr.Status = item.Status;
        formProps.tr.EndUseId = item.EndUseId;
        formProps.tr.WorkTypeId = item.WorkTypeId;
        formProps.tr.Title = item.Title;


      });


    batch.execute().then((value) => {
      this.reactElement = React.createElement(TrForm, formProps);
      ReactDom.render(this.reactElement, this.domElement);
    }
    );

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
                PropertyPaneTextField('trListUrl', {
                  label: strings.ListUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected loadData() {
    let batch = pnp.sp.createBatch();
    pnp.sp.web.lists.getByTitle("End Uses").items.inBatch(batch).get()
      .then((items) => {
        let newProps: ITrFormProps = _.clone(this.reactElement.props);
        newProps.endUses = _.map(items, (item) => {
          return new EndUse(item["Id"], item["Title"], item["ApplicationTypeId"]);
        });
        this.reactElement.props = newProps;

      });
    pnp.sp.web.lists.getByTitle("Work Types").items.inBatch(batch).get()
      .then((items) => {

        this.reactElement.props.workTypes = _.map(items, (item) => {
          return new WorkType(item["Id"], item["Title"]);
        });

      });
    pnp.sp.web.lists.getByTitle("Application Types").inBatch(batch).items.get()
      .then((items) => {

        this.reactElement.props.applicationTypes = _.map(items, (item) => {
          return new ApplicationType(item["Id"], item["Title"], item["WorkTypeId"]);
        });

      });
    batch.execute().then((value) => {

      console.log("All done!");
    }
    );
  }
  private save(tr: TR): Promise<any> {
   
    return pnp.sp.web.lists.getByTitle("trs").items.getById(tr.Id).update(tr);
    // .then((results) => {
    //   debugger;
    // })
    // .catch((reaseon) => {
    //   debugger;
    // });
  }
}

