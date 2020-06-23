declare interface IDataVerticalsWebPartStrings {
  General: {
    WebPartDefaultTitle: string;
  },
  PropertyPane: {
    DataVerticalsGroupName: string;
    Verticals: {
      PropertyLabel: string;
      PanelHeader: string;
      PanelDescription: string;
      ButtonLabel: string;
      Fields: {
        TabName: string;
        IconName: string;
        IsLink: string;
        LinkUrl: string;
        OpenBehavior: string;
      }
    }
  }
}

declare module 'DataVerticalsWebPartStrings' {
  const strings: IDataVerticalsWebPartStrings;
  export = strings;
}
