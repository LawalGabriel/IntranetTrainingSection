declare interface IIntranetTrainingsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListTitleFieldLabel: string;
}

const strings: IIntranetTrainingsStrings = {
  PropertyPaneDescription: 'Configure your training items web part',
  BasicGroupName: 'Basic Settings',
  DescriptionFieldLabel: 'Description',
  ListTitleFieldLabel: 'List Title'
};

declare module 'IntranetTrainingsWebPartStrings' {
  const strings: IIntranetTrainingsWebPartStrings;
  export = strings;
}
