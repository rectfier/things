import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IAttachment {
  name: string;
  content: ArrayBuffer | Blob;
  category?: string;
  status?: string;
}

export interface IDocumentServiceParams {
  sp: SPFI;
  projectId: string;
  attachments: IAttachment[];
  list: string;
  path: string;
}

export const uploadDocumentsToDocSet = async (params: IDocumentServiceParams): Promise<void> => {
  const { sp, projectId, attachments, list, path } = params;

  // External Web
  const url = "url here"; // Replace with actual URL
  const extWeb = Web([sp.web, url]);

  // Get Library info with select and expand
  const extLib = extWeb.lists.getByTitle(list);
  const libInfo = await extLib.select("Title", "RootFolder/ServerRelativeUrl").expand("RootFolder")();

  // Build libRootUrl and docSetPath
  const libRootUrl = libInfo.RootFolder.ServerRelativeUrl;
  const docSetPath = `${libRootUrl}/${path}/${projectId}`;

  // Check if doc set exists
  let exists = false;
  try {
    await extWeb.getFolderByServerRelativePath(docSetPath)();
    exists = true;
  } catch {
    exists = false;
  }

  if (!exists) {
    // Create the folder using the library root path
    const libRootFolder = extWeb.getFolderByServerRelativePath(libRootUrl);
    await libRootFolder.addSubFolderUsingPath(`${path}/${projectId}`);

    // Get the created folder's list item and set Title and ContentTypeId
    const newFolder = extWeb.getFolderByServerRelativePath(docSetPath);
    const folderItem = await newFolder.listItemAllFields();

    await extLib.items.getById(folderItem.Id).update({
      Title: projectId,
      ContentTypeId: "0x0120D520" // Document Set content type ID
    });
  }

  // Get folder reference for uploads
  const extFolderRef = extWeb.getFolderByServerRelativePath(docSetPath);

  // Upload attachments and update metadata on each file
  for (const attachment of attachments) {
    const fileResult = await extFolderRef.files.addUsingPath(attachment.name, attachment.content, { Overwrite: true });

    // Update file metadata (category and status)
    const fileItem = await extWeb.getFileByServerRelativePath(fileResult.ServerRelativeUrl).listItemAllFields();
    await extLib.items.getById(fileItem.Id).update({
      DocumentCategory: attachment.category || "",
      ProjectStatus: attachment.status || ""
    });
  }
};
