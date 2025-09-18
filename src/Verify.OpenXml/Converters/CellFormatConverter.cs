class CellFormatConverter :
    WriteOnlyJsonConverter<CellFormat>
{
    public override void Write(VerifyJsonWriter writer, CellFormat options)
    {
        writer.WriteStartObject();
        writer.WriteMember(options, options.Alignment, "Alignment");
        writer.WriteMember(options, options.Protection, "Protection");
        writer.WriteMember(options, options.FontId, "FontId");
        writer.WriteMember(options, options.FillId, "FillId");
        writer.WriteMember(options, options.BorderId, "BorderId");

        if (options.NumberFormatId != null)
        {
            var format = NumberFormatConverter.GetReadableFormat(options.NumberFormatId, VerifyOpenXml.CurrentDocument!.WorkbookPart!);
            writer.WriteMember(options, format, "NumberFormatId");
        }

        writer.WriteMember(options, options.FormatId, "FormatId");
        writer.WriteMember(options, options.ApplyNumberFormat, "ApplyNumberFormat");
        writer.WriteMember(options, options.ApplyFont, "ApplyFont");
        writer.WriteMember(options, options.ApplyFill, "ApplyFill");
        writer.WriteMember(options, options.ApplyBorder, "ApplyBorder");
        writer.WriteMember(options, options.ApplyAlignment, "ApplyAlignment");
        writer.WriteMember(options, options.ApplyProtection, "ApplyProtection");
        writer.WriteMember(options, options.PivotButton, "PivotButton");
        writer.WriteMember(options, options.QuotePrefix, "QuotePrefix");
        writer.WriteEndObject();
    }
}