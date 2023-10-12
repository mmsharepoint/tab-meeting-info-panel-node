import { AzureNamedKeyCredential, TableClient } from "@azure/data-tables";
import { ICustomer } from "../../model/ICustomer";

export const getCustomer = async (meetingID: string): Promise<ICustomer> => {
    const tableClient = getAZTableClient();
    const customerEntities = await tableClient.listEntities({ disableTypeConversion: false, queryOptions: { filter: `PartitionKey eq '${meetingID}'`}})
    const customerEntity = await customerEntities.next();
    const customer: ICustomer = {
        Name: customerEntity.value.Name,
        Email: customerEntity.value.Email,
        Phone: customerEntity.value.Phone,
        Id: customerEntity.value.rowKey
    }
    return customer;
}

export const getAZTableClient = (): TableClient => {
    const accountName: string = process.env.AZURE_TABLE_ACCOUNTNAME!;
    const storageAccountKey: string = process.env.AZURE_TABLE_KEY!;
    const storageUrl = `https://${accountName}.table.core.windows.net/`;
    const tableClient = new TableClient(storageUrl, "Customer", new AzureNamedKeyCredential(accountName, storageAccountKey));
    return tableClient;
}