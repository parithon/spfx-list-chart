export interface ISPWeb {
  Url: string;
  Title: string;
  IsRootWeb?: boolean;
  Lists?: ISPList[];
}

export interface ISPList {
  Id: string;
  Title: string;
  BaseTemplate?: number;
}
