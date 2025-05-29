import axios from 'axios';
import { IInputs } from "../../generated/ManifestTypes"; // adjust path if needed

export interface IAttachment {
    attachmentId: string;
    body: string; // Base64 encoded attachment content
    filename: string;
    filesize: number;
    mimetype: string;
}

export interface IPartyList {
    partyid: string; name: string; entitytype: string;
}

export interface IRecord {
    id: string;
    subject: string;
    from: IPartyList[];
    sendermailboxidname: string;
    to: IPartyList[];
    torecipients: string;
    regardingobjectidname: string;
    regardingobjectid: string;
    priority: string;
    status: string;
    description?: string;
    ownerid?: string;
    owneridname?: string;
    cc: IPartyList[];
    bcc: IPartyList[];
    createdon: string;
    createdby: string;
    createdbyname: string;
    directioncode: string;
    activitytypecode: string;
    scheduledend: string;
    scheduledstart: string;
    senton: string;
    attachments?: IAttachment[];
    [key: string]: string | number | boolean | null | undefined | IAttachment[] | IPartyList[];
}



export let emailData: IRecord[] = [];

// Fetch all email records
export async function fetchAllEmailData(context: ComponentFramework.Context<IInputs>): Promise<IRecord[]> {
    try {
        // let apikey;
        // await context.webAPI.retrieveRecord("strava_apikey", "ed8577ba-5e3c-f011-b4cc-6045bdff083f", "?$select=strava_key").then(
        //     function success(result) {
        //         apikey = result["strava_key"];
        //     },
        //     function (error) {
        //         console.log(error.message);
        //     }
        // );
        // if (!apikey) {
        //     throw new Error("API key not found.");
        //     return [];
        // }
        const url = "https://ccd-eus-dev1-wa-dataanalytics-archival.azurewebsites.net/api/DataAnalyticsArchival/EmailDetails/91B01626-9D14-EC11-B6E6-00224822E38C";
        const headers = {
            "authentication_key": "45c18dd37644161267e534b31f1f7186933d4bec21f62b1c3ca75a97f6f4a2b1",
        };

        const response = await axios.get(url, { headers });
        const dataArray = response.data;

        if (!Array.isArray(dataArray)) {
            throw new Error("API did not return an array.");
        }

        emailData = dataArray.map(mapToIRecord);
        return emailData;

    } catch (error) {
        console.error("Error fetching email data:", error);
        throw error;
    }
}

function mapToIRecord(data: any): IRecord {
    const attachments = (data.attachments?.files || []).map((file: any, index: number): IAttachment => ({
        attachmentId: `attach-${index + 1}`,
        body: file.base64Content ?? "", // <-- confirm actual field name
        filename: file.fileName ?? `file-${index + 1}`,
        filesize: file.fileSize ?? 0,
        mimetype: file.mimeType ?? "application/octet-stream",
    }));
    return {
        id: data.id,
        activityid: data.activityid,
        statecode: data.statecode,
        status: "Draft",
        statuscode: data.statuscode,
        prioritycode: data.prioritycode,
        priority: "High",
        directioncode: "Outgoing",
        createdby: data.createdby,
        createdby_entitytype: data.createdby_Entitytype,
        createdbyname: data.createdbyname,
        modifiedby: data.modifiedby,
        modifiedby_entitytype: data.modifiedby_Entitytype,
        owningbusinessunit: data.owningbusinessunit,
        owningbusinessunit_entitytype: data.owningbusinessunit_Entitytype,
        owninguser: data.owninguser,
        owninguser_entitytype: data.owninguser_Entitytype,
        regardingobjectid: data.regardingobjectid,
        regardingobjectid_entitytype: data.regardingobjectid_Entitytype,
        sendermailboxid: data.sendermailboxid,
        sendermailboxid_entitytype: data.sendermailboxid_Entitytype,
        ownerid: data.ownerid,
        ownerid_entitytype: data.ownerid_Entitytype,
        from: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" }
        ],
        to: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "f60407dd-9bd6-ea11-a813-000d3a1bb158", name: "MARCIA JOHNSON", entitytype: "contact" },
            { partyid: "ede79e36-ced5-ea11-a813-000d3a1bb158", name: "MICHAEL FAULK", entitytype: "contact" }
        ],
        cc: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "f60407dd-9bd6-ea11-a813-000d3a1bb158", name: "MARCIA JOHNSON", entitytype: "contact" },
            { partyid: "ede79e36-ced5-ea11-a813-000d3a1bb158", name: "MICHAEL FAULK", entitytype: "contact" }
        ],
        bcc: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "f60407dd-9bd6-ea11-a813-000d3a1bb158", name: "MARCIA JOHNSON", entitytype: "contact" },
            { partyid: "ede79e36-ced5-ea11-a813-000d3a1bb158", name: "MICHAEL FAULK", entitytype: "contact" }
        ],
        activitytypecode: data.activitytypecode,
        attachmentcount: data.attachmentcount,
        createdon: data.createdon,
        descriptionblobid: data.descriptionblobid,
        descriptionblobid_name: data.descriptionblobid_name,
        messageid: data.messageid || "",
        modifiedbyname: data.modifiedbyname,
        modifiedon: data.modifiedon,
        owneridname: data.owneridname,
        owneridyominame: data.owneridyominame,
        owningbusinessunitname: data.owningbusinessunitname,
        regardingobjectidname: data.regardingobjectidname,
        sender: data.sender || "",
        sendermailboxidname: data.sendermailboxidname || "",
        senton: data.senton,
        subject: data.subject || "",
        torecipients: data.torecipients || "",
        scheduledend: data.scheduledend,
        scheduledstart: data.scheduledstart,
        description: data.description || "",
        attachments,
    };
}
