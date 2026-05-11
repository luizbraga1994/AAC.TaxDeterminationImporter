namespace AAC.TaxDeterminationImporter.Core.DAO
{
    internal static class HanaFunctions
    {
        internal const string FN_TAXCODE_EXPORT =
            @"CREATE OR REPLACE FUNCTION FN_TAXCODE_EXPORT(keyType INT, keyValue NVARCHAR(200))
RETURNS description NVARCHAR(1000)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER AS
BEGIN
    IF :keyType IN (1, 2, 6, 16, 17) THEN
        description := :keyValue;
    ELSEIF :keyType = 3 THEN
        description := (SELECT ""Descrip"" FROM OMGP WHERE TO_VARCHAR(""AbsEntry"") = :keyValue);
    ELSEIF :keyType = 5 THEN
        description := (SELECT ""NcmCode"" FROM ONCM WHERE TO_VARCHAR(""AbsEntry"") = :keyValue);
    ELSEIF :keyType = 9 THEN
        description := (SELECT ""ItmsGrpNam"" FROM OITB WHERE TO_VARCHAR(""ItmsGrpCod"") = :keyValue);
    ELSEIF :keyType IN (11, 12) THEN
        description := (SELECT ""GroupName"" FROM OCRG WHERE TO_VARCHAR(""GroupCode"") = :keyValue);
    ELSE
        description := :keyValue;
    END IF;
END;";

        internal const string FN_GET_KEYFIELDDESCRIPTION =
            @"CREATE OR REPLACE FUNCTION FN_GET_KEYFIELDDESCRIPTION(keyType INT, keyValue NVARCHAR(200))
RETURNS description NVARCHAR(1000)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER AS
BEGIN
    IF :keyType IN (1, 2, 6, 17) THEN
        description := :keyValue;
    ELSEIF :keyType = 3 THEN
        description := (SELECT ""Descrip"" FROM OMGP WHERE TO_VARCHAR(""AbsEntry"") = :keyValue);
    ELSEIF :keyType = 5 THEN
        description := (SELECT ""NcmCode"" FROM ONCM WHERE TO_VARCHAR(""AbsEntry"") = :keyValue);
    ELSEIF :keyType = 9 THEN
        description := (SELECT ""ItmsGrpNam"" FROM OITB WHERE TO_VARCHAR(""ItmsGrpCod"") = :keyValue);
    ELSEIF :keyType IN (11, 12) THEN
        description := (SELECT ""GroupName"" FROM OCRG WHERE TO_VARCHAR(""GroupCode"") = :keyValue);
    ELSEIF :keyType = 16 THEN
        description := (SELECT ""BPLName"" FROM OBPL WHERE TO_VARCHAR(""BPLId"") = :keyValue);
    ELSE
        description := :keyValue;
    END IF;
END;";
    }
}

