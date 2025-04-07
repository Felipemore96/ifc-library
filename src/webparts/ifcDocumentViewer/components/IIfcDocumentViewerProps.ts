import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentLibraryViewerProps {
  title: string;
  description: string;
  context: WebPartContext;
  libraryName: string; // The name of the document library to display
  displayMode: number;
  updateProperty: (value: string) => void;
}
