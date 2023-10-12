// Initialize dotenv, to use .env file settings if existing
require("dotenv").config();
const azure = require('@azure/identity');
const authProviders = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const appConfigClient = require('@azure/app-configuration');
const { TableClient, AzureNamedKeyCredential } = require("@azure/data-tables");

const graph = require('@microsoft/microsoft-graph-client');

const main = async () => {
    const userPrincipalName = process.env.MEETING_OWNER;
    const dummyAttendee = process.env.MEETING_ATTENDEE;

    const meetingSubject = "Test Meeting with App / Tab Node";

    const customerName = "Contoso";
    const customerEmail = "JohnJohnson@contoso.com";
    const customerPhone = "+491515445556";
    const customerId = "47110815";

    console.log(`Creating meeting with Owner ${userPrincipalName} nad Subject ${meetingSubject}`);

    // @azure/identity
    const credential = new azure.ClientSecretCredential(
        process.env.TENANT_ID,
        process.env.MICROSOFT_APP_ID,
        process.env.MICROSOFT_APP_PASSWORD
    );
    
    // @microsoft/microsoft-graph-client/authProviders/azureTokenCredentials
    const authProvider = new authProviders.TokenCredentialAuthenticationProvider(credential, {
        // The client credentials flow requires that you request the
        // /.default scope, and pre-configure your permissions on the
        // app registration in Azure. An administrator must grant consent
        // to those permissions beforehand.
        scopes: ['https://graph.microsoft.com/.default'],
    });

    const graphClient = graph.Client.initWithMiddleware({ authProvider: authProvider });

    // credential.getToken([
    //     'https://graph.microsoft.com/.default'
    // ]).then(result => {
    //     console.log(`Àccess Token: ${result.token}`);
    // });

    const userId = await getUserId(graphClient, userPrincipalName);
    const joinUrl = await createEvent(graphClient, meetingSubject, userPrincipalName, userId, dummyAttendee);
    // console.log(joinUrl);
    const chatId = await getMeetingChatId(graphClient, userId, joinUrl)    
    console.log(chatId);
    const appId = await getAppId(graphClient);
    console.log(appId);
    await installAppInChat(graphClient, appId, chatId);
    await installTabInChat(graphClient, appId, chatId);
    const customer = {
        Name: customerName,
        Phone: customerPhone,
        Email: customerEmail,
        Id: customerId
    }
    saveConfig(chatId, customer);
    await createCustomer(chatId, customer);
    const checkCustomer = await getCustomer(chatId);
    console.log('Customer created in Azure Table');
    console.log(checkCustomer);
};

async function getUserId (client, upn) {
    const user = await client.api(`/users/${upn}`)
	.get();
    return user.id;
};

async function createEvent (client, meetingSubject, userPrincipalName, userID, dummyAttendee) {
    const startDate = new Date();
    const endDate = new Date(startDate);
    const startHours = startDate.getHours();
    endDate.setHours(startHours + 1);
    console.log(startDate.toISOString());
    console.log(endDate.toISOString());
    const event = {
        subject: meetingSubject,
        isOnlineMeeting: true,
        start: {
            dateTime: startDate.toISOString(),
            timeZone: 'Europe/Berlin'
        },
        end: {
            dateTime: endDate.toISOString(),
            timeZone: 'Europe/Berlin'
        },
        Organizer :
            {
                emailAddress: { address: userPrincipalName }
            }, 
        attendees: [
        {
            emailAddress: {
            address: dummyAttendee
            },
            type: 'required'
        }
        ]
    };
    
    const result = await client.api(`/users/${userID}/events`)
        .post(event);
    return result.onlineMeeting.joinUrl;
}

async function getMeetingChatId (client, userID, joinUrl) {
    const onlineMeeting = await client.api(`/users/${userID}/onlineMeetings`)
        .filter(`joinWebUrl eq '${joinUrl}'`)
        .get();    
    const chatId = onlineMeeting.value[0].chatInfo.threadId;
    console.log(`OnlineMeeting with ChatID ${chatId}`);
    return chatId;
}

async function getAppId (client) {
    const apps = await client.api('/appCatalogs/teamsApps')                
                .filter("distributionMethod eq 'organization' and displayName eq 'Teams Meeting Custom Data'")
                .get();

    let appId = "";
    console.log(apps);
    console.log(apps.value);
    appId = apps.value[0].id;

    return appId;
}

async function installAppInChat (client, appId, chatId) {
    const requestBody =
    {
        "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${appId}`
    };
    await client.api(`/chats/${chatId}/installedApps`)
            .post(requestBody);
    return true;
}

async function installTabInChat(client, appId, chatId) {
    const teamsTab = {
        displayName: 'Custom Data',
        'teamsApp@odata.bind': `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${appId}`,
        configuration: {
            entityId: "2DCA2E6C7A10415CAF6B8AB6661B3154", // ToDo
            contentUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetingDataTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`,
            removeUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetingDataTab/remove.html?theme={theme}`,
            websiteUrl: `https://${process.env.PUBLIC_HOSTNAME}/meetingDataTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
        }
      };
      
      await client.api(`/chats/${chatId}/tabs`)
          .post(teamsTab);
}

async function saveConfig (meetingID, newConfig) {
    const client = getAppConfigClient();
    if (newConfig.Name) {
      await client.setConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERNAME`, value: newConfig.Name });
    }
    if (newConfig.Phone) {
      await client.setConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERPHONE`, value: newConfig.Phone });
    }
    if (newConfig.Email) {
      await client.setConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMEREMAIL`, value: newConfig.Email });
    }
    if (newConfig.Id) {
      await client.setConfigurationSetting({ key: `TEAMSMEETINGSERVICECALL:${meetingID}:CUSTOMERID`, value: newConfig.Id });
    }
}

function getAppConfigClient () {
    const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING;
    let client;
    if (connectionString.startsWith('Endpoint=')) {
      client = new appConfigClient.AppConfigurationClient(connectionString);      
    }
    else {
      const credential = new azure.DefaultAzureCredential();
      client = new appConfigClient.AppConfigurationClient(connectionString, credential);
    }  
    return client;
}

async function createCustomer(meetingID, customer)
{
    const tableClient = getAZTableClient();
    const tableEntity = 
    {
        partitionKey: meetingID,
        rowKey: customer.Id,
        Name: customer.Name,
        Email: customer.Email,
        Phone: customer.Phone
    };
   await tableClient.createEntity(tableEntity);
}

function getAZTableClient () {
    const accountName = process.env.AZURE_TABLE_ACCOUNTNAME;
    const storageAccountKey = process.env.AZURE_TABLE_KEY;
    const storageUrl = `https://${accountName}.table.core.windows.net/`;
    const tableClient = new TableClient(storageUrl, "Customer", new AzureNamedKeyCredential(accountName, storageAccountKey));
    return tableClient;
}

async function getCustomer (meetingID) {
    const tableClient = getAZTableClient();   
    const customerEntities = await tableClient.listEntities({ disableTypeConversion: false, queryOptions: { filter: `PartitionKey eq '${meetingID}'`}})
    const customerEntity = await customerEntities.next();
    const customer = {
        Name: customerEntity.value.Name,
        Email: customerEntity.value.Email,
        Phone: customerEntity.value.Phone,
        Id: customerEntity.value.rowKey
    }
    return customer;
}

main();

