import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./IfcDocumentViewer.module.scss";
import { IDocumentLibraryViewerProps } from "./IIfcDocumentViewerProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  CommandBar,
  ICommandBarItemProps,
} from "@fluentui/react/lib/CommandBar";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";
import { Text } from "@fluentui/react/lib/Text";
import { IconButton } from "@fluentui/react/lib/Button";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";

export interface IDocument {
  Id: string;
  Title: string;
  Name: string;
  ModifiedBy: string;
  Modified: string;
  FileRef: string;
  FileType: string;
}

export interface IDocumentLibraryViewerState {
  documents: IDocument[];
  isLoading: boolean;
  error: string | null;
}

const DocumentLibraryViewer: React.FC<IDocumentLibraryViewerProps> = (
  props
) => {
  const [state, setState] = useState<IDocumentLibraryViewerState>({
    documents: [],
    isLoading: true,
    error: null,
  });

  const getDocuments = (): void => {
    const { context, libraryName } = props;
    const libraryToUse = libraryName || "Documents"; // Default to Documents if not specified

    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libraryToUse}')/items?$select=Id,Title,FileLeafRef,Modified,FileRef,Editor/Title&$expand=Editor&$orderby=Modified desc`;

    context.spHttpClient
      .get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Error fetching documents: ${response.statusText}`);
        }
      })
      .then((data) => {
        if (data && data.value) {
          const docs = data.value.map((item: any) => ({
            Id: item.Id,
            Title: item.Title || "No Title",
            Name: item.FileLeafRef,
            ModifiedBy: item.Editor ? item.Editor.Title : "Unknown",
            Modified: new Date(item.Modified).toLocaleDateString(),
            FileRef: item.FileRef,
            FileType: item.FileLeafRef.split(".").pop().toLowerCase(),
          }));

          setState({
            documents: docs,
            isLoading: false,
            error: null,
          });
        }
      })
      .catch((error) => {
        setState({
          documents: [],
          isLoading: false,
          error: error.message,
        });
        console.error("Error fetching documents:", error);
      });
  };

  useEffect(() => {
    getDocuments();
  }, []);

  const openDocument = (documentUrl: string): void => {
    window.open(documentUrl, "_blank");
  };

  const handleCustomAction = (): void => {
    // Implement your custom action here
    alert(
      "Custom action triggered! You can implement your specific functionality here."
    );
    // For example: Open a modal, trigger a flow, etc.
  };

  const commandItems: ICommandBarItemProps[] = [
    {
      key: "refresh",
      text: "Refresh",
      iconProps: { iconName: "Refresh" },
      onClick: getDocuments,
    },
    {
      key: "customAction",
      text: "Custom Action",
      iconProps: { iconName: "CustomList" },
      onClick: handleCustomAction,
    },
  ];

  const columns: IColumn[] = [
    {
      key: "name",
      name: "Name",
      fieldName: "Name",
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IDocument) => {
        return (
          <a href="#" onClick={() => openDocument(item.FileRef)}>
            {item.Name}
          </a>
        );
      },
    },
    {
      key: "title",
      name: "Title",
      fieldName: "Title",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "modifiedBy",
      name: "Modified By",
      fieldName: "ModifiedBy",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 70,
      maxWidth: 70,
      onRender: (item: IDocument) => {
        const menuProps: IContextualMenuProps = {
          items: [
            {
              key: "open",
              text: "Open",
              iconProps: { iconName: "OpenFile" },
              onClick: () => openDocument(item.FileRef),
            },
            {
              key: "customAction",
              text: "Custom Action",
              iconProps: { iconName: "CustomList" },
              onClick: () => handleCustomAction(),
            },
          ],
        };

        return (
          <IconButton
            menuProps={menuProps}
            iconProps={{ iconName: "MoreVertical" }}
          />
        );
      },
    },
  ];

  const { isLoading, error, documents } = state;

  return (
    <div className={styles.documentLibraryViewer}>
      <div className={styles.header}>
        <Text variant="large">{props.title || "Document Library"}</Text>
      </div>

      <CommandBar items={commandItems} />

      {isLoading && <div>Loading documents...</div>}
      {error && <div className={styles.error}>Error: {error}</div>}

      {!isLoading && !error && (
        <DetailsList
          items={documents}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
        />
      )}
    </div>
  );
};

export default DocumentLibraryViewer;
