declare interface IUnpromoteNewsCommandSetStrings {
  Command1: string;
  Doyouwant:string;
  AlreadyNews:string;
  WIP:string;
  Newsispromoted:string;
}

declare module 'UnpromoteNewsCommandSetStrings' {
  const strings: IUnpromoteNewsCommandSetStrings;
  export = strings;
}
