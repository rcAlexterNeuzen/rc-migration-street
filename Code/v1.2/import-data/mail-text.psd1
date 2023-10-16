@{
  Default   = "<html><head><style>table, th, td {
            border: 1px solid black;
            border-collapse: collapse;
          }
          th, td {
            padding: 5px;
          }
          th {
            text-align: left;
          }
          </style><body>"
  Fileshare = @(
    @{
      Success = "Hi,<br>
            <br>
            The migration of fileshares is completed.<br>
            <br>
            All jobs are completed.<br>
            Attached you will see the results<br>
            <br>
            Kind Regards,<br>
            <br>
            Rapid Circle Migration Street</body>"
      Failed  = "Hi,<br>
            <br>
            The migration of fileshares is failed.<br>
            <br>
            Attached you will see the results<br>
            <br>
            Kind Regards,<br>
            <br>
            Rapid Circle Migration Street</body>"
    }
  )
  Homedir = @(
    @{
      Success = "Hi,<br>
            <br>
            The migration of homedirs is completed.<br>
            <br>
            All jobs are completed.<br>
            Attached you will see the results<br>
            <br>
            Kind Regards,<br>
            <br>
            Rapid Circle Migration Street</body>"
      Failed  = "Hi,<br>
            <br>
            The migration of homedirs is failed.<br>
            <br>
            Attached you will see the results<br>
            <br>
            Kind Regards,<br>
            <br>
            Rapid Circle Migration Street</body>"
    }
  )
}
