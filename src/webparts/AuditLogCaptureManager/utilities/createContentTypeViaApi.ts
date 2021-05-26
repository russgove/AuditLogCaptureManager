import { IAuditLogCaptureManagerProps } from '../components/IAuditLogCaptureManagerProps';
import { callManagementApi } from './callManagementApi';
// export async function createContentType(siteUrl: string): Promise<string> {
// debugger
//   const context: SP.ClientContext = new SP.ClientContext(decodeURIComponent(siteUrl));
//   var itemContentType = context.get_site().get_rootWeb().get_contentTypes().getById("0x01");
//   context.load(itemContentType);
//   await executeQuery(context)
//     .catch((err) => {
//       console.log(err);
//       debugger;
//     });


//   var contentTypeCreationInformation = new SP.ContentTypeCreationInformation();
//   contentTypeCreationInformation.set_name("Audit Item");
//   contentTypeCreationInformation.set_description("Microsoft 365 SharePoint Audit Capture detail record");
//   contentTypeCreationInformation.set_parentContentType(itemContentType);

//   var newContentType: SP.ContentType = context.get_site().get_rootWeb().get_contentTypes()
//     .add(contentTypeCreationInformation);
//   await addFields(context, newContentType);
//   await executeQuery(context)
//     .catch((err) => {
//       console.log(err);
//     });
//   return newContentType.get_stringId();
// }
export async function createContentType(siteUrl: string, parentContext: IAuditLogCaptureManagerProps) {

  const url = `${parentContext.managementApiUrl}/api/AddContentTypeToSite?siteUrl=${encodeURIComponent(siteUrl)}`;
  await callManagementApi(parentContext.aadHttpClient, url, "POST");
  alert(`ct created`);
}

async function addTextField(context: SP.ClientContext, newContentType: SP.ContentType, fieldName: string, displayName: string, description: string): Promise<any> {
  var fieldXML = getTextFieldSchema(fieldName, description, displayName);
  var field = context.get_site().get_rootWeb().get_fields().addFieldAsXml(fieldXML, true, SP.AddFieldOptions.addFieldToDefaultView);
  context.load(field);
  var fieldLink = new SP.FieldLinkCreationInformation();
  fieldLink.set_field(field);
  fieldLink.get_field().set_required(false);
  fieldLink.get_field().set_hidden(false);
  newContentType.get_fieldLinks().add(fieldLink);
  newContentType.update(false);
  return;
}
function getTextFieldSchema(fieldName: string, displayName: string, description: string): string {
  return `<Field Type="Text" Name="${fieldName}" DisplayName="${description}" Required="FALSE" Group="_Audit Columns" />`;
}

async function addNumberField(context: SP.ClientContext, newContentType: SP.ContentType, fieldName: string, displayName: string, description: string): Promise<any> {
  var fieldXML = getNumberFieldSchema(fieldName, description, displayName);
  var field = context.get_site().get_rootWeb().get_fields().addFieldAsXml(fieldXML, true, SP.AddFieldOptions.addFieldToDefaultView);
  var fieldLink = new SP.FieldLinkCreationInformation();
  fieldLink.set_field(field);
  fieldLink.get_field().set_required(false);
  fieldLink.get_field().set_hidden(false);
  newContentType.get_fieldLinks().add(fieldLink);
  return;
}

function getNumberFieldSchema(fieldName: string, displayName: string, description: string): string {
  return `<Field Type="Number" Name="${fieldName}" DisplayName="${description}" Required="FALSE" Group="_Audit Columns" />`;
}

function executeQuery(context: SP.ClientContext): Promise<any> {


  const promise: Promise<any> = new Promise<any>((resolve, reject) => {
    try {
      context.executeQueryAsync(
        (sender: any, args: SP.ClientRequestSucceededEventArgs) => {
   
          return resolve(args);
        },
        (sender: any, err: SP.ClientRequestFailedEventArgs) => {
          debugger;
          alert(err.get_message());
          console.log(err.get_errorDetails());
          return reject(err.get_message());
        }
      );
    }
    catch (err) {
    
      console.log(err);
      debugger;

    }

  });
  return promise;
}