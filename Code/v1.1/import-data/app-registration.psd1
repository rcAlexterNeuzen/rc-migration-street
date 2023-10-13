# graph api applications permissions
@{
    GraphApi = "00000003-0000-0000-c000-000000000000"
    AppRegistration = @(
        @{
            Name           = "Rapid Circle Migration Street"
            Description    = "App registration for migration street Rapid Circle"
            SignInAudience = "AzureADMyOrg"
            web            = "https://portal.azure.com"
        }
    )
    perm            = @(
        "f3a65bd4-b703-46df-8f7e-0174fea562aa" # Channel.Create (Application)
        "59a6b24b-4225-4393-8165-ebaec5f55d7a" # Channel.ReadBasic.All (Application)
        "35930dcf-aceb-4bd1-b99a-8ffed403c974" # ChannelMember.ReadWrite.All (Application)
        "c97b873f-f59f-49aa-8a0e-52b32d762124" # ChannelSettings.Read.All (Application)
        "332a536c-c7ef-4017-ab91-336970924f0d" # Sites.Read.All (Application)
        "23fc2474-f741-46ce-8465-674744c5c361" # Team.Create (Application)
        "2280dda6-0bfd-44ee-a2f4-cb867cfc4c1e" # Team.ReadBasic.All (Application)
        "0121dc95-1b9f-4aed-8bac-58c5ac466691" # TeamMember.ReadWrite.All (Application)
        "62a82d76-70ea-41e2-9197-370581804d09" # Group.ReadWrite.All (Application)
        "a82116e5-55eb-4c41-a434-62fe8a61c773" # Sites.FullControl.All (Application)
        "75359482-378d-4052-8f01-80520e7db3cd" # Files.ReadWrite.All (application)
        "b633e1c5-b582-4048-a93e-9f11b44c7e96" # Mail.Send (application)
        "df021288-bdef-4463-88db-98f22de89214" # User.Read.All (application)
    )
}
# reference id from : https://learn.microsoft.com/en-us/graph/permissions-reference#all-permissions-and-ids

