if count >= 13
    for i=firstIndex:lastIndex
        Operand = TheMFE.GetOperandAt(i);
        if strcmp(Operand.Type, 'CENY') == true
            Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
            spot_setting = Spot1.GetSettings();
            spot_setting.Field.SetFieldNumber(i);
            FieldCENY = Operand.GetCellAt(4);
            Analyse{n, 7} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);               
            n = n+1;
            Spot1.Close;
        end
    end
end    