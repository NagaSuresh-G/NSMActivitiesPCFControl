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
    to: IPartyList[];
    regardingobjectidname: string;
    priority: string;
    status: string;
    description?: string;
    ownerid?: string;
    createdon: string;
    createdby: string;
    createdbyname: string;
    activitytypecode: string;
    scheduledend: string;
    attachments?: IAttachment[];
    [key: string]: string | number | boolean | null | undefined | IAttachment[] | IPartyList[];
}

export const phoneCallData: IRecord[] = [
    {
        id: "phone_001",
        activityid: "DA171177-E140-EC11-8C62-000D3A57FAA2",
        subject: "Follow-up on Client Meeting",
        createdon: "5/10/2025, 10:00:00 AM",
        createdby: "ba60d8ef-19eb-4443-81de-546385f9409f",
        createdby_entitytype: "systemuser",
        createdbyname: "Strava Dynamics 365 Admin",
        activitytypecode: "PhoneCall",
        statecode: 1,
        status: "Draft",
        statuscode: 3,
        prioritycode: 1,
        priority: "High",
        regardingobjectid: "518DE740-7E3E-EC11-B6E5-00224821155B",
        regardingobjectid_entitytype: "strava_policy",
        regardingobjectidname: "CSK04A00190209",
        from: [
            { partyid: "a2515b12-7e01-ef11-a1fd-000d3a10e128", name: "Haley Phillips", entitytype: "systemuser" }
        ],
        to: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "17372dce-4a83-ee11-8179-000d3a9bc712", name: "SMITH", entitytype: "account" },
            { partyid: "d436c7fb-9e9e-ef11-8a6a-6045bdda0302", name: "AARON ABEL", entitytype: "contact" }
        ],
        to_entitytype: "systemuser",
        toname: "Paul Choi",
        scheduledend: "5/10/2025, 10:30:00 AM",
        phonenumber: "+1234567890",
        actualstart: "5/10/2025, 10:00:00 AM",
        actualend: "5/10/2025, 10:30:00 AM",
        directioncode: "Outgoing",
        ownerid: "a2515b12-7e01-ef11-a1fd-000d3a10e128",
        ownerid_entitytype: "systemuser",
        owneridname: "Haley Phillips",
        new_phonecallreason: "Follow-up",
        new_phonecalloutcome: "Completed",
        strava_note: "Discussed project milestones and next steps.",
        strava_notecategory: "Client Follow-up",
        strava_publishnote: "No",
        description: "Ensure all action items from the meeting are addressed and client feedback is incorporated.",
        softphon_queuename: "Sales Queue",
        softphon_queuetimeseconds: 120,
        softphon_genesyscloudwrapupcode: "Follow-up",
        softphon_ivrtimeseconds: 60,
        actualdurationminutes: "30",
        softphon_durationseconds: 1800,
        softphon_dispositiondurationseconds: 120,
        softphon_interactionurl: "https://example.com/interaction/12345",
    },
    {
        id: "phone_002",
        activityid: "DA171177-E140-EC11-8C62-000D3A57FAA0",
        subject: "Test Phone Call",
        createdon: "5/10/2025, 10:00:00 AM",
        createdby: "ba60d8ef-19eb-4443-81de-546385f9409f",
        createdby_entitytype: "systemuser",
        createdbyname: "Strava Dynamics 365 Admin",
        activitytypecode: "PhoneCall",
        statecode: 1,
        status: "Draft",
        statuscode: 3,
        prioritycode: 1,
        priority: "High",
        regardingobjectid: "518DE740-7E3E-EC11-B6E5-00224821155B",
        regardingobjectid_entitytype: "strava_policy",
        regardingobjectidname: "CSK04A00190209",
         from: [
            { partyid: "a2515b12-7e01-ef11-a1fd-000d3a10e128", name: "Haley Phillips", entitytype: "systemuser" }
        ],
        from_entitytype: "systemuser",
        fromname: "Haley Phillips",
        to: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "17372dce-4a83-ee11-8179-000d3a9bc712", name: "SMITH", entitytype: "account" },
            { partyid: "d436c7fb-9e9e-ef11-8a6a-6045bdda0302", name: "AARON ABEL", entitytype: "contact" }
        ],
        to_entitytype: "systemuser",
        toname: "Paul Choi",
        scheduledend: "5/10/2025, 10:30:00 AM",
        phonenumber: "+1234567890",
        actualstart: "5/10/2025, 10:00:00 AM",
        actualend: "5/10/2025, 10:30:00 AM",
        directioncode: "Outgoing",
        ownerid: "a2515b12-7e01-ef11-a1fd-000d3a10e128",
        ownerid_entitytype: "systemuser",
        owneridname: "Haley Phillips",
        new_phonecallreason: "Follow-up",
        new_phonecalloutcome: "Completed",
        strava_note: "Discussed project milestones and next steps.",
        strava_notecategory: "Client Follow-up",
        strava_publishnote: "No",
        description: "Ensure all action items from the meeting are addressed and client feedback is incorporated.",
        softphon_queuename: "Sales Queue",
        softphon_queuetimeseconds: 120,
        softphon_genesyscloudwrapupcode: "Follow-up",
        softphon_ivrtimeseconds: 60,
        actualdurationminutes: "30",
        softphon_durationseconds: 1800,
        softphon_dispositiondurationseconds: 120,
        softphon_interactionurl: "https://example.com/interaction/12345",
    },
    {
        id: "phone_003",
        activityid: "DA171177-E140-EC11-8C62-000D3A57FAA3",
        subject: "Follow up PCF control development",
        createdon: "5/10/2025, 10:00:00 AM",
        createdby: "ba60d8ef-19eb-4443-81de-546385f9409f",
        createdby_entitytype: "systemuser",
        createdbyname: "Strava Dynamics 365 Admin",
        activitytypecode: "PhoneCall",
        statecode: 1,
        status: "Draft",
        statuscode: 3,
        prioritycode: 1,
        priority: "High",
        regardingobjectid: "518DE740-7E3E-EC11-B6E5-00224821155B",
        regardingobjectid_entitytype: "strava_policy",
        regardingobjectidname: "CSK04A00190209",
         from: [
            { partyid: "a2515b12-7e01-ef11-a1fd-000d3a10e128", name: "Haley Phillips", entitytype: "systemuser" }
        ],
        from_entitytype: "systemuser",
        fromname: "Haley Phillips",
        to: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "17372dce-4a83-ee11-8179-000d3a9bc712", name: "SMITH", entitytype: "account" },
            { partyid: "d436c7fb-9e9e-ef11-8a6a-6045bdda0302", name: "AARON ABEL", entitytype: "contact" }
        ],
        to_entitytype: "systemuser",
        toname: "Paul Choi",
        scheduledend: "5/10/2025, 10:30:00 AM",
        phonenumber: "+1234567890",
        actualstart: "5/10/2025, 10:00:00 AM",
        actualend: "5/10/2025, 10:30:00 AM",
        directioncode: "Outgoing",
        ownerid: "a2515b12-7e01-ef11-a1fd-000d3a10e128",
        ownerid_entitytype: "systemuser",
        owneridname: "Haley Phillips",
        new_phonecallreason: "Follow-up",
        new_phonecalloutcome: "Completed",
        strava_note: "Discussed project milestones and next steps.",
        strava_notecategory: "Client Follow-up",
        strava_publishnote: "No",
        description: "Ensure all action items from the meeting are addressed and client feedback is incorporated.",
        softphon_queuename: "Sales Queue",
        softphon_queuetimeseconds: 120,
        softphon_genesyscloudwrapupcode: "Follow-up",
        softphon_ivrtimeseconds: 60,
        actualdurationminutes: "30",
        softphon_durationseconds: 1800,
        softphon_dispositiondurationseconds: 120,
        softphon_interactionurl: "https://example.com/interaction/12345",
    },
    {
        id: "phone_004",
        activityid: "DA171177-E140-EC11-8C62-000D3A57FAA4",
        subject: "Follow up PCF control development 4",
        createdon: "5/10/2025, 10:00:00 AM",
        createdby: "ba60d8ef-19eb-4443-81de-546385f9409f",
        createdby_entitytype: "systemuser",
        createdbyname: "Strava Dynamics 365 Admin",
        activitytypecode: "PhoneCall",
        statecode: 1,
        status: "Draft",
        statuscode: 3,
        prioritycode: 1,
        priority: "High",
        regardingobjectid: "518DE740-7E3E-EC11-B6E5-00224821155B",
        regardingobjectid_entitytype: "strava_policy",
        regardingobjectidname: "CSK04A00190209",
         from: [
            { partyid: "a2515b12-7e01-ef11-a1fd-000d3a10e128", name: "Haley Phillips", entitytype: "systemuser" }
        ],
        from_entitytype: "systemuser",
        fromname: "Haley Phillips",
        to: [
            { partyid: "e2fe7e1e-5253-ea11-a816-000d3a579c8c", name: "Paul Choi", entitytype: "systemuser" },
            { partyid: "17372dce-4a83-ee11-8179-000d3a9bc712", name: "SMITH", entitytype: "account" },
            { partyid: "d436c7fb-9e9e-ef11-8a6a-6045bdda0302", name: "AARON ABEL", entitytype: "contact" }
        ],
        to_entitytype: "systemuser",
        toname: "Paul Choi",
        scheduledend: "5/10/2025, 10:30:00 AM",
        phonenumber: "+1234567890",
        actualstart: "5/10/2025, 10:00:00 AM",
        actualend: "5/10/2025, 10:30:00 AM",
        directioncode: "Outgoing",
        ownerid: "a2515b12-7e01-ef11-a1fd-000d3a10e128",
        ownerid_entitytype: "systemuser",
        owneridname: "Haley Phillips",
        new_phonecallreason: "Follow-up",
        new_phonecalloutcome: "Completed",
        strava_note: "Discussed project milestones and next steps.",
        strava_notecategory: "Client Follow-up",
        strava_publishnote: "No",
        description: "Ensure all action items from the meeting are addressed and client feedback is incorporated.",
        softphon_queuename: "Sales Queue",
        softphon_queuetimeseconds: 120,
        softphon_genesyscloudwrapupcode: "Follow-up",
        softphon_ivrtimeseconds: 60,
        actualdurationminutes: "30",
        softphon_durationseconds: 1800,
        softphon_dispositiondurationseconds: 120,
        softphon_interactionurl: "https://example.com/interaction/12345",
    },
];