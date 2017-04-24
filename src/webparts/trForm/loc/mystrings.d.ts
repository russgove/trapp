declare interface ITrFormStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ModeFieldLabel: string;
}

declare module 'trFormStrings' {
  const strings: ITrFormStrings;
  export = strings;
}
