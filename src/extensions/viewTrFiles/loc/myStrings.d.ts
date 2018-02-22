declare interface IViewTrFilesCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ViewTrFilesCommandSetStrings' {
  const strings: IViewTrFilesCommandSetStrings;
  export = strings;
}
