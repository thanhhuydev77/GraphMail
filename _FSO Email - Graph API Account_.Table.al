table 70000 "Email - Graph API Account"
{
    DataCaptionFields = "Email Address";
    DataClassification = CustomerContent;

    fields
    {
        field(70000; ID; Guid)
        {
            Caption = 'ID';
            DataClassification = SystemMetadata;
        }
        field(70010; "Email Address"; Text[250])
        {
            Caption = 'Email Address';
            DataClassification = CustomerContent;
        }
        field(70020; Name; Text[250])
        {
            Caption = 'Account Name';
            DataClassification = CustomerContent;
        }
        field(70030; "Graph APIEmail Connector"; Enum "Email Connector")
        {
            DataClassification = SystemMetadata;
        }
        field(70040; "Tenant ID"; Guid)
        {
            Caption = 'Tenant ID';
            DataClassification = CustomerContent;
        }
        field(70050; "Client ID"; Guid)
        {
            Caption = 'Client ID';
            DataClassification = CustomerContent;
        }
        field(70060; "Client Secrect"; Text[250])
        {
            Caption = 'Client Secrect';
            DataClassification = CustomerContent;
            extendedDatatype = Masked;
        }
    }
    keys
    {
        key(PK; "ID")
        {
            Clustered = true;
        }
        key(Idx01; "Email Address")
        {
            Unique = true;
        }
    }
}
