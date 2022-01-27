declare interface ICustomButtonCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CustomButtonCommandSetStrings' {
  const strings: ICustomButtonCommandSetStrings;
  export = strings;
}
