@{
    HomeDir = 'Homefolder'
    SharepointSite = ''
    Errors = 'Errorlist'
    FileShare = 'Fileshare'
    Ids = @{
        FileShareChannelId = ''
        HomeFolderChannelid = ''
        ErrorChannelId = ''
        GeneralChannelId = ''
        TeamsId = ''
    }
    Mails = @{
        FileshareChannel = ''
        HomedriveChannel = ''
        ErrorChannel = ''
        GeneralChannel = ''
    }
    folders = @{
        FileshareFolder = "fileshare%20migrations/"
        HomeFolderFolder = "homefolder%20migrations/"
        ErrorFolder = "error%20migrations/"
        GeneralFolder = "/" 
    }
    TeamsName = 'Rapid Circle Migrations'
    Channels = @(
        'Fileshare Migrations'
        'Homefolder Migrations'
        'Errors'
    )
}
