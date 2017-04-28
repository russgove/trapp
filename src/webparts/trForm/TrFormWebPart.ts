import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from "sp-pnp-js";
import { SearchQuery, SearchResults, SortDirection } from "sp-pnp-js";
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration, PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import { TR, WorkType, ApplicationType, EndUse, modes, User } from "./dataModel";
import * as strings from 'trFormStrings';
import * as _ from 'lodash';
import TrForm from './components/TrForm';
import { ITrFormProps } from './components/ITrFormProps';
import { ITrFormWebPartProps } from './ITrFormWebPartProps';
import {
   IPersonaProps, PersonaPresence
} from 'office-ui-fabric-react';

export default class TrFormWebPart extends BaseClientSideWebPart<ITrFormWebPartProps> {
  private tr: TR;
  private childTRs: Array<TR>;



  private reactElement: React.ReactElement<ITrFormProps>;
  private trContentTypeID: string;
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
    if (document.getElementById("s4-ribbonrow")) {
      document.getElementById("s4-ribbonrow").style.display = "none";
    }

    let formProps: ITrFormProps = {
      childTRs: [],
      techSpecs: [],
      requestors: [],
      cancel: this.cancel.bind(this),
      ensureUser: this.ensureUser,
      mode: this.properties.mode,
      TRsearch: this.TRsearch.bind(this),
      peoplesearch: this.peoplesearch,
      workTypes: [],
      applicationTypes: [],
      endUses: [],
      tr: new TR(),
      save: this.save.bind(this)
    };
    let batch = pnp.sp.createBatch();
    // get the Technincal Request content type so we can use it later in searches
    pnp.sp.web.contentTypes.inBatch(batch).get()
      .then((contentTypes) => {
        debugger;
        const trContentTyoe = _.find(contentTypes, (contentType) => { return contentType["Name"] === "TechnicalRequest" ;});
        this.trContentTypeID = trContentTyoe["Id"]["StringValue"];
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching content types'");
        console.log(error.message);

      });

    pnp.sp.web.siteGroups.getByName("TR YY Tech Specialists").users.inBatch(batch).get()
      .then((items) => {
        formProps.techSpecs = _.map(items, (item) => {
          return new User(item["Id"], item["Title"], item["Title"], item["Title"]);
        });
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching Tech Specialists'");
        console.log(error.message);

      });
    pnp.sp.web.siteGroups.getByName("TR YY Requestors").users.inBatch(batch).get()
      .then((items) => {
        formProps.requestors = _.map(items, (item) => {
          return new User(item["Id"], item["Title"], item["Title"], item["Title"]);
        });
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching Requestors'");
        console.log(error.message);

      });
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
    var queryParameters = new UrlQueryParameterCollection(window.location.href);

    if (this.properties.mode !== modes.NEW) {
      if (queryParameters.getValue("Id")) {
        const id: number = parseInt(queryParameters.getValue("Id"));
        let fields = "*,ParentTR/Title,Requestor/Title";
        // get the requested tr
        pnp.sp.web.lists.getByTitle("Technical Requests").items.getById(id).expand("ParentTR,Requestor").select(fields).inBatch(batch).get()

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
            debugger;
            formProps.tr.TechSpecId = item.TechSpecId;

          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem");
            console.log(error.message);

          });
        // get the Child trs
        const self = this;
        pnp.sp.web.lists.getByTitle("Technical Requests").items.filter("ParentTR eq " + id).expand("ParentTR,Requestor").select(fields).inBatch(batch).get()
          .then((items) => {
            // this may resilve befor we get the mainn tr, so jyst stash them away for now.
            for (const item of items) {
              let childtr: TR = new TR();
              childtr.Id = item.Id;
              childtr.ActualCompletionDate = item.ActualCompletionDate;
              childtr.ApplicationTypeId = item.ApplicationTypeId;
              childtr.ActualStartDate = item.ActualStartDate;
              childtr.CER = item.CER;
              childtr.Customer = item.Customer;
              childtr.TRDueDate = item.TRDueDate;
              childtr.EstimatedHours = item.EstimatedHours;
              childtr.InitiationDate = item.InitiationDate;
              childtr.TRPriority = item.TRPriority;
              childtr.RequestorId = item.RequestorId;
              if (item.Requestor) {
                childtr.RequestorName = item.Requestor.Title;
              }
              childtr.Site = item.Site;
              childtr.Status = item.Status;
              childtr.EndUseId = item.EndUseId;
              childtr.WorkTypeId = item.WorkTypeId;
              childtr.Title = item.Title;
              childtr.TitleArea = item.TitleArea;
              childtr.DescriptionArea = item.DescriptionArea;
              childtr.SummaryArea = item.SummaryArea;
              childtr.ParentTRId = item.ParentTRId;
              if (item.ParentTR) {
                childtr.ParentTR = item.ParentTR.Title;
              }
              debugger;
              childtr.TechSpecId = item.TechSpecId;
              formProps.childTRs.push(childtr);
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

    let queryText = "{0} Path:{1}* ContentTypeId:{2}*";
    queryText = queryText
      .replace("{0}", searchText)
      .replace("{1}", this.context.pageContext.web.absoluteUrl)
      .replace("{2}", this.trContentTypeID);
    let sq: SearchQuery = {
      Querytext: queryText,
      RowLimit: 50,
      SelectProperties: ["Title", "MailBoxOWSTEXT", "ListItemID", "CEROWSTEXT", "CustomerOWSTEXT", "SiteOWSTEXT"]
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      // SelectProperties: ["*"]
    };
    console.log(sq);

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


  private navigateToSource() {
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    let encodedSource = queryParameters.getValue("Source");
    if (encodedSource) {
      let source = decodeURIComponent(encodedSource);
      console.log("source uis " + source);
      window.location.href = source;
    }
  }
  private save(tr: TR): Promise<any> {
    // remove lookups
    let copy = _.clone(tr) as any;
    delete copy.RequestorName;
    delete copy.ParentTR;
    let technicalSpecialists = (copy.TechSpecId) ? copy.TechSpecId : [];
    delete copy.TechSpecId;
    copy["TechSpecId"] = {};
    copy["TechSpecId"]["results"] = technicalSpecialists;




    if (tr.Id !== null) {
      return pnp.sp.web.lists.getByTitle("Technical Requests").items.getById(tr.Id).update(copy).then((x) => {

        this.navigateToSource();
      });
    }
    else {
      return pnp.sp.web.lists.getByTitle("Technical Requests").items.add(copy).then((x) => {

        this.navigateToSource();

      });
    }

  }
  private cancel(): void {
    this.navigateToSource();

  }
}

