using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// Verifies that a slide's text and its rendered image land on the same page. A page is numbered by
// where its slide sits in p:sldIdLst, the order the deck is shown in, which here is not the order
// the slide parts were added in: read in part order, page 1 would carry the text of the third slide.
public class PowerpointPagesTests
{
    #region PagesToInclude

    [Test]
    public async Task PagesToInclude()
    {
        using var presentation = ThreeSlides();
        await Verify(presentation)
            .PagesToInclude(2);
    }

    #endregion

    // sample.pptx, so that the master, layout and theme PowerPoint requires are there. Its own
    // slide becomes the third, and the two added after it the first and second.
    static PresentationDocument ThreeSlides()
    {
        // Copied, since a stream over the bytes could not grow to hold the added slides
        var stream = new MemoryStream();
        using (var file = File.OpenRead("sample.pptx"))
        {
            file.CopyTo(stream);
        }

        var presentation = PresentationDocument.Open(stream, true);
        var presentationPart = presentation.PresentationPart!;
        var slideIds = presentationPart.Presentation!.SlideIdList!;
        var sample = presentationPart.SlideParts.Single();

        AddSlide("First", 257);
        AddSlide("Second", 258);

        SetText(sample.Slide!, "Third");
        var sampleId = slideIds.GetFirstChild<SlideId>()!;
        sampleId.Remove();
        slideIds.Append(sampleId);

        return presentation;

        void AddSlide(string text, uint id)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = (Slide)sample.Slide!.CloneNode(true);
            SetText(slidePart.Slide, text);
            slidePart.AddPart(sample.SlideLayoutPart!);
            slideIds.Append(
                new SlideId
                {
                    Id = id,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });
        }

        static void SetText(Slide slide, string text) =>
            slide.Descendants<A.Text>().Single().Text = text;
    }
}
