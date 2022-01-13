declare interface ICopyExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CopyExtensionCommandSetStrings' {
  const strings: ICopyExtensionCommandSetStrings;
  export = strings;
}
