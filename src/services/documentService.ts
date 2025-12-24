import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/lists";
import "@pnp/sp/content-types";

export interface IAttachment {
  name: string;
  content: ArrayBuffer | Blob;
  DocumentCategory?: string;
  ProjectStatus?: string;
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

  // Construct the external URL
  const url = "url here"; // Replace with actual URL
  const extWeb = Web([sp.web, url]);

  // Document Set name is the projectId
  // Path structure: List/path/projectId (Document Set)/files
  const docSetPath = `${list}/${path}/${projectId}`;

  // Check if document set exists
  let docSetExists = false;
  try {
    await extWeb.getFolderByServerRelativePath(docSetPath)();
    docSetExists = true;
  } catch (e) {
    docSetExists = false;
  }

  // If it doesn't exist, create the Document Set
  if (!docSetExists) {
    // Create the Document Set folder (recursively creates path if needed)
    // using addUsingPath handles both checking and creating
    await extWeb.folders.addUsingPath(docSetPath);
    
    // Get the Document Set content type from the list
    const listObj = extWeb.lists.getByTitle(list);
    
    // In PnPjs v4, contentTypes is a sub-collection accessed via the .contentTypes property
    const contentTypes = await listObj.contentTypes();
    const docSetContentType = contentTypes.find(ct => ct.Name === "Document Set");

    if (!docSetContentType) {
      throw new Error("Document Set content type not found on the list.");
    }
    
    // Update the folder to use Document Set content type
    // We fetch the Item from the folder we just ensured exists
    const docSetFolder = extWeb.getFolderByServerRelativePath(docSetPath);
    const docSetItem = await docSetFolder.listItemAllFields();
    
    // Use metadata from first attachment if available
    const firstAttachment = attachments.length > 0 ? attachments[0] : null;

    // Ensure we are updating the content type and metadata correctly
    await listObj.items.getById(docSetItem.Id).update({
      ContentTypeId: docSetContentType.StringId,
      DocumentCategory: firstAttachment?.DocumentCategory || "",
      ProjectStatus: firstAttachment?.ProjectStatus || ""
    });
  } else if (attachments.length > 0) {
    // If it exists, update metadata from the latest batch
    const listObj = extWeb.lists.getByTitle(list);
    const docSetFolder = extWeb.getFolderByServerRelativePath(docSetPath);
    const docSetItem = await docSetFolder.listItemAllFields();
    const firstAttachment = attachments[0];

    await listObj.items.getById(docSetItem.Id).update({
      DocumentCategory: firstAttachment.DocumentCategory || "",
      ProjectStatus: firstAttachment.ProjectStatus || ""
    });
  }

  // Upload attachments to the document set
  for (const attachment of attachments) {
    await extWeb.getFolderByServerRelativePath(docSetPath).files.addUsingPath(attachment.name, attachment.content, { Overwrite: true });
  }
};
