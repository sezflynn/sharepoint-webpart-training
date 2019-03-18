declare interface IHelloArthurWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloArthurWebPartStrings' {
  const strings: IHelloArthurWebPartStrings;
  export = strings;
}
