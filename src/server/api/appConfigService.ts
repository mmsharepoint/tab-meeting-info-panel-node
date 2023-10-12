import { AppConfigurationClient } from "@azure/app-configuration";
import { DefaultAzureCredential } from "@azure/identity";
import { ICustomer } from "../../model/ICustomer";

export const retrieveConfig = async (meetingID: string): Promise<ICustomer> => {
  const client = getClient();
  let customerName  = "";
  let customerPhone  = "";
  let customerEmail  = "";
  let customerId  = "";
  try {
    const movie1Setting = await client.getConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERNAME`});
    customerName = movie1Setting.value!;
    const movie2Setting = await client.getConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERPHONE`});
    customerPhone = movie2Setting.value!;
    const movie3Setting = await client.getConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMEREMAIL`});
    customerEmail = movie3Setting.value!;
    const containerIdSetting = await client.getConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERID`});
    customerId = containerIdSetting.value!;
  }
  catch(error) {
    
  }
  return Promise.resolve({ Name: customerName, Phone: customerPhone, Email: customerEmail, Id: customerId });
};
    
const getClient = (): AppConfigurationClient => {
  const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
  let client: AppConfigurationClient;
  if (connectionString.startsWith('Endpoint=')) {
    client = new AppConfigurationClient(connectionString);
  }
  else {
    const credential = new DefaultAzureCredential();
    client = new AppConfigurationClient(connectionString, credential);
  }  
  return client;
};