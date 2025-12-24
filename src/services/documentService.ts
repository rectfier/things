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
    // Get the Document Set content type from the list
    const listObj = extWeb.lists.getByTitle(list);
    const contentTypes = await listObj.contentTypes();
    const docSetContentType = contentTypes.find(ct => ct.Name === "Document Set");

    if (!docSetContentType) {
      throw new Error("Document Set content type not found on the list.");
    }

    // Create the Document Set
    const parentFolderPath = `${list}/${path}`;
    await extWeb.getFolderByServerRelativePath(parentFolderPath).addSubFolderUsingPath(projectId);
    
    // Update the folder to use Document Set content type
    const docSetFolder = extWeb.getFolderByServerRelativePath(docSetPath);
    const docSetItem = await docSetFolder.listItemAllFields();
    await extWeb.lists.getByTitle(list).items.getById(docSetItem.Id).update({
      ContentTypeId: docSetContentType.StringId
    });
  }

  // Upload attachments to the document set
  for (const attachment of attachments) {
    await extWeb.getFolderByServerRelativePath(docSetPath).files.addUsingPath(attachment.name, attachment.content, { Overwrite: true });
  }
};
