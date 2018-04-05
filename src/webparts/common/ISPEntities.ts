import Guid from "@microsoft/sp-core-library/lib/Guid";

export interface ISPList {
  Id: Guid;
  Title: string;
  BaseTemplate?: number;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPField {
  Id: Guid;
  ListId: Guid;
  Title: string;
  InternalName?: string;
  TypeAsString?: string;
}

export interface ISPFields {
  value: ISPField[];
}
