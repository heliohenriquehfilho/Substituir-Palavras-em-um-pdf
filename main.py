import fitz

def adjust_matrix(font, bbox, text, fontsize):
    """Compute matrix performing a horizontal scale.

    Args:
        font: Font object
        bbox: bbox of the text to fill
        text: text
        fontsize: fontsize to use
    Returns:
        Horizontal scaling matrix
    """
    tl = font.text_length(text, fontsize=fontsize)
    width = bbox[2] - bbox[0]
    scale = width / tl
    return fitz.Matrix(scale, 1)


def get_fontlist(page):
    """Make a dictionary for existing fonts."""
    flist = {}
    for f in page.get_fonts():
        # extract xref, full font name and reference name
        xref, fullname, refname = f[0], f[3], f[4]

        # subset fonts have a "+" in string position 7
        fullname = fullname[6:] if "+" in fullname else fullname
        ff = doc.extract_font(xref)  # extract font buffer
        font = fitz.Font(fontbuffer=ff[-1])

        # store reference name and the Font object
        flist[font.name] = (refname, font)
    return flist


doc = fitz.open(r"C:\Users\hh100801\Downloads\relatorio_premissas.pdf")
page = doc[0]
page.clean_contents()  # page needs cleaning for correct positions of inserts

# make a dictionary of fonts used on this page
fontlist = get_fontlist(page)
print("Font list:", fontlist)  # Debug: Print the font list

# we intend to replace all occurrences of "[demanda_pta]" by "5,98 MW"
bboxes = page.search_for("[demanda_pta]")

spans = []  # store occurrences here
for bbox in bboxes:  # extract full text meta info for each occurrence
    for b in page.get_text("dict", clip=bbox)["blocks"]:
        for l in b["lines"]:
            spans.extend(l["spans"])

# Debug: Print the spans
print("Spans:", spans)

# now redact away the word "passenger"
for s in spans:
    page.add_redact_annot(s["bbox"])
    if s["font"] == "1":
        s["font"] = "1 Regular"

    fontlist['1 Regular'] = ("F1", fitz.Font("figo"))
        
    print(s)

page.apply_redactions(images=0, graphics=0, text=0)

# now insert new text in emptied bboxes
for s in spans:
    point = fitz.Point(s["origin"])  # insertion point
    text = s["text"]  # original text ([demanda_pta])
    fsize = s["size"]  # fontsize
    font = s["font"]  # font name in PDF

    # Debug: Print the font being used
    print("Processing span with font:", font)

    # extract-convert color - note we use red for demo-purposes
    color = fitz.sRGB_to_pdf(s["color"])  # re-use original color
    # replace old by new text
    text = text.replace("[demanda_pta]", "12.8798 kW")

    # choose right font for output:
    # there often exists ambiguity WRT "-" instead of spaces etc. so we
    # make a second try when encountering problems

    try:
        fontname, font_obj = fontlist[font]
    except KeyError:
        try:
            fontname, font_obj = fontlist[font.replace("-", " ")]
        except KeyError:
            # Debug: Print the error and continue
            print(f"Font '{font}' not found in font list. Skipping span.")
            continue

    # IMPORTANT: matrix to stretch or shrink new text horizontally
    matrix = adjust_matrix(font_obj, s["bbox"], text, fsize)
    print(f"Inserting text: '{text}' at point: {point} with font: {fontname} and size: {fsize}")
    page.insert_text(
        point,
        text,
        fontname=fontname,
        fontsize=fsize,
        color=color,
        morph=(point, matrix),  # this will shrink / stretch horizontally
    )

doc.subset_fonts()  # not needed in this version: reusing existing font subsets
doc.ez_save(r"C:\Users\hh100801\Documents\PDFs\trains-new2.pdf")
