import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from "sp-pnp-js";
import { SearchQuery, SearchResults, SortDirection } from "sp-pnp-js";
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField, PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { TR, WorkType, ApplicationType, EndUse, modes } from "./dataModel";
import * as strings from 'trFormStrings';
import * as _ from 'lodash';
import TrForm from './components/TrForm';
import { ITrFormProps } from './components/ITrFormProps';
import { ITrFormWebPartProps } from './ITrFormWebPartProps';
import {
  Dropdown, IDropdownProps, IPersonaProps, PersonaPresence
} from 'office-ui-fabric-react';

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
// hide the ribbon

   document.getElementById("s4-ribbonrow").style.display = "none";

    let formProps: ITrFormProps = { ensureUser: this.ensureUser, mode: this.properties.mode, TRsearch: this.TRsearch, peoplesearch: this.peoplesearch, workTypes: [], applicationTypes: [], endUses: [], tr: new TR(), save: this.save };
    let batch = pnp.sp.createBatch();
    pnp.sp.web.lists.getByTitle("End Uses").items.inBatch(batch).get()
      .then((items) => {
        formProps.endUses = _.map(items, (item) => {
          return new EndUse(item["Id"], item["Title"], item["ApplicationTypeId"]);
        });
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'End uses'");
        console.log(error.message);

      });
    pnp.sp.web.lists.getByTitle("Work Types").items.inBatch(batch).get()
      .then((items) => {
        formProps.workTypes = _.map(items, (item) => {
          return new WorkType(item["Id"], item["Title"]);
        });

      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'Work Types'");
        console.log(error.message);

      });
    pnp.sp.web.lists.getByTitle("Application Types").items.inBatch(batch).get()
      .then((items) => {

        formProps.applicationTypes = _.map(items, (item) => {
          return new ApplicationType(item["Id"], item["Title"], item["WorkTypesId"]);
        });

      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'Application Types'");
        console.log(error.message);

      });
    var queryParameters= new UrlQueryParameterCollection(window.location.href);
    debugger;
    if (this.properties.mode !== modes.NEW) {
      if (queryParameters.getValue("Id")) {
        const id: number = parseInt(queryParameters.getValue("Id"));
        let fields = "*,ParentTR/Title,Requestor/Title";
        pnp.sp.web.lists.getByTitle("trs").items.getById(id).expand("ParentTR,Requestor").select(fields).inBatch(batch).get()
          .then((item) => {
            debugger;
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
            formProps.tr.RequestorId = item.RequestorId;
            if (item.Requestor) {
              formProps.tr.RequestorName = item.Requestor.Title;
            }
            formProps.tr.Site = item.Site;
            formProps.tr.Status = item.Status;
            formProps.tr.EndUseId = item.EndUseId;
            formProps.tr.WorkTypeId = item.WorkTypeId;
            formProps.tr.Title = item.Title;
            formProps.tr.TitleArea = item.TitleArea;
            formProps.tr.DescriptionArea = item.DescriptionArea;
            formProps.tr.SummaryArea = item.SummaryArea;
            formProps.tr.ParentTRId = item.ParentTRId;
            if (item.ParentTR) {
              formProps.tr.ParentTR = item.ParentTR.Title;
            }

          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem");
            console.log(error.message);

          });
      }
    }
    else {
      console.log("ERROR, Id not specified with New or Edit form");
    }

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
                PropertyPaneDropdown('mode', {
                  label: strings.ModeFieldLabel,
                  options: [
                    { text: "New", key: modes.NEW },
                    { text: "Edit", key: modes.EDIT },
                    { text: "Display", key: modes.DISPLAY },
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
  public TRsearch(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps[]> {


    let sq: SearchQuery = {
      Querytext: searchText + " Path:https://tronoxglobal.sharepoint.com/sites/TR/TRs ContentTypeId:0x0100A87C11965AD4494DA4E5CB2CC25622BB",
      RowLimit: 50,
      SelectProperties: ["Title", "MailBoxOWSTEXT", "ListItemID", "CEROWSTEXT", "CustomerOWSTEXT", "SiteOWSTEXT"]
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      // SelectProperties: ["*"]
    };
    console.log(sq);
    debugger;
    return pnp.sp.search(sq).then((results: SearchResults) => {
      let resultsPersonas: Array<IPersonaProps> = [];
      for (let element of results.PrimarySearchResults) {
        const temp = element as any;
        let personapprop: IPersonaProps = {
          primaryText: temp.Title,
          secondaryText: temp.CustomerOWSTEXT,
          tertiaryText: temp.SiteOWSTEXT,
          presence: PersonaPresence.none,
          id: temp.ListItemID
        };
        resultsPersonas.push(personapprop);
      }
      return _.sortBy(resultsPersonas, "primaryText");
    });


  }
  public peoplesearch(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps[]> {


    let sq: SearchQuery = {
      Querytext: "Title:" + searchText + "* ContentClass=urn:content-class:SPSPeople",
      SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",  //http://www.dotnetmafia.com/blogs/dotnettipoftheday/archive/2013/01/04/list-of-common-sharepoint-2013-result-source-ids.aspx
      RowLimit: 50,
      SelectProperties: ["PreferredName", "Department", "JobTitle", "PictureURL",
        "OfficeNumber", "WorkEmail"]
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      // SelectProperties: ["*"]
    };
    return pnp.sp.search(sq).then((results: SearchResults) => {
      let resultsPersonas: Array<IPersonaProps> = [];
      for (let element of results.PrimarySearchResults) {
        const temp = element as any;
        let personapprop: IPersonaProps = {
          primaryText: temp.PreferredName,
          secondaryText: temp.JobTitle,
          tertiaryText: (temp.OfficeNumber) ? temp.Department + "(" + temp.OfficeNumber + ") " : temp.Department,
          imageUrl: temp.PictureURL,
          imageInitials: temp.contentclass,
          presence: PersonaPresence.none,
          optionalText: temp.WorkEmail // need this for ensureuser

        };
        resultsPersonas.push(personapprop);
      }
      return _.sortBy(resultsPersonas, "primaryText");
    });


  }
  protected ensureUser(email): Promise<any> {
    return pnp.sp.web.ensureUser(email);
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
    // remove lookups
    let copy = _.clone(tr);
    delete copy.RequestorName;
    delete copy.ParentTR;

    if (tr.Id !== -1) {
      return pnp.sp.web.lists.getByTitle("trs").items.getById(tr.Id).update(copy);
    }
    else {
      return pnp.sp.web.lists.getByTitle("trs").items.add(copy);
    }
    // .then((results) => {
    //   debugger;
    // })
    // .catch((reaseon) => {
    //   debugger;
    // });
  }
}

