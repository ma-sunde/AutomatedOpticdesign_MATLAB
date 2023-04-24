% Convert file to Non-sequential mode
    convertNSmode = TheSystem.Tools.OpenConvertToNSCGroup();
    convertNSmode.ConvertFileToNSC = true;
    convertNSmode.RunAndWaitForCompletion();
    convertNSmode.Close();
