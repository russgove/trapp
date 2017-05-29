declare interface ITrTimeCardStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'trTimeCardStrings' {
  const strings: ITrTimeCardStrings;
  export = strings;
}
