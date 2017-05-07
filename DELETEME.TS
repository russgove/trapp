 public render(): void {
  let formProps: ITrFormProps = {
      workTypes: [],
      tr:new TR()
     };
    let batch = pnp.sp.createBatch();
   pnp.sp.web.lists.getByTitle(this.properties.workTypeListName).items.inBatch(batch).get()
      .then((items) => {
        formProps.workTypes = _.map(items, (item) => {
          return new WorkType(item["Id"], item["Title"]);
        });

      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching 'Work Types' from list named " + this.properties.workTypeListName);
        console.log(error.message);

      });
     var queryParameters = new UrlQueryParameterCollection(window.location.href);

    if (this.properties.mode !== modes.NEW) {
      if (queryParameters.getValue("Id")) {
        const id: number = parseInt(queryParameters.getValue("Id"));
        let fields = "*,ParentTR/Title,Requestor/Title";
        let expands = "ParentTR,Requestor";
        // get the requested tr
        pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(id).expand(expands).select(fields).inBatch(batch).get()

          .then((item) => {
            formProps.tr = new TR();
            this.moveFieldsToTR(formProps.tr, item);


          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem  from list named " + this.properties.technicalRequestListName);;
            console.log(error.message);

          });
      }
      else {
        console.log("ERROR, Id not specified with Display or Edit form");
      }
    }


    batch.execute().then((value) => {

      this.reactElement = React.createElement(TrForm, formProps);
      var formComponent: TrForm = ReactDom.render(this.reactElement, this.domElement) as TrForm;
 
    }
    );

  }