declare interface ITrFormStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListUrlFieldLabel: string;
}

declare module 'trFormStrings' {
  const strings: ITrFormStrings;
  export = strings;
}
