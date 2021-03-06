declare interface IExternalFieldUpdaterCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExternalFieldUpdaterCommandSetStrings' {
  const strings: IExternalFieldUpdaterCommandSetStrings;
  export = strings;
}
