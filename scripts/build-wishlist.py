"""Regenerate the static gift list from the PowerPoint. Python 3, standard library only."""
from pathlib import Path
import html
import json
import re
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "dependencies" / "wishlist 2026 ari.pptx"
PAGE = ROOT / "wishlist.html"
# Sizes supplied by the family for the webpage, without editing the source deck.
WEB_SIZE_OVERRIDES = {(4, "Ashlyn striped top"): "XXS"}
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
STORES = {
    "us.brandymelville.com": "Brandy Melville",
    "www.hollisterco.com": "Hollister",
    "www.tillys.com": "Tillys",
    "www.chicknlegs.com": "ChicknLegs",
    "www.ugg.com": "UGG",
    "www.stevemadden.com": "Steve Madden",
    "www.baublebar.com": "BaubleBar",
    "www.victoriassecret.com": "Victoria’s Secret",
    "www.walmart.com": "Walmart",
    "www.freestyleusa.com": "Freestyle",
    "www.ever-eden.com": "Evereden",
    "www.thehydrojug.com": "HydroJug",
    "www.petermillar.com": "Peter Millar",
    "www.amazon.com": "Amazon",
}


def clean(value):
    return re.sub(r"[\u200b-\u200d\ufeff]", "", value).replace("\xa0", " ").strip()


def paragraphs(element):
    lines = []
    for paragraph in element.findall(".//a:p", NS):
        value = "".join(node.text or "" if node.tag == f"{{{NS['a']}}}t" else "\n"
                        for node in paragraph.iter()
                        if node.tag in (f"{{{NS['a']}}}t", f"{{{NS['a']}}}br"))
        lines.extend(clean(line) for line in value.split("\n"))
    return lines


def product_name(element):
    value = " ".join(p for p in paragraphs(element) if p)
    return re.split(r"\s*(?:price|website|size|color)\s*:", value, flags=re.I)[0].strip()


def product_size(elements):
    # Use the product's own shapes/cells so multi-product slides keep their
    # sizes separate. Preserve the author's wording, including pant length.
    text = "\n".join(p for element in elements for p in paragraphs(element))
    match = re.search(r"^[ \t]*size[ \t]*:[ \t]*([^\n]+)", text, re.I | re.M)
    return clean(match[1]) if match else None


def display_size(value):
    # Normalize webpage labels only; retain the original extracted source sizes.
    value = re.sub(r"\b(?:xss|xxs)\b", "XXS", value, flags=re.I)
    return re.sub(r"\bWomen(?:['’]s)?\s+5\.5\b", "Women's 5.5", value, flags=re.I)


def product_url(value):
    value = re.sub(r"\s+", "", clean(value))
    parsed = urllib.parse.urlsplit(value)
    # Preserve the actual destination already supplied by the author, rather
    # than making the family travel through an advertising redirect.
    if parsed.hostname == "go.skimresources.com":
        value = urllib.parse.parse_qs(parsed.query).get("url", [value])[0]
    elif parsed.hostname == "ad.doubleclick.net" and parsed.query.startswith("https://"):
        value = parsed.query
    parsed = urllib.parse.urlsplit(value)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise ValueError(f"Invalid source product URL: {value!r}")
    return value


def links(elements, relationships):
    found = []
    for element in elements:
        # Concatenate text runs before parsing: slide 20 splits "https" and "://".
        text = "\n".join(paragraphs(element))
        text = re.sub(r"https?\s*:\s*//", lambda m: re.sub(r"\s+", "", m[0]), text)
        found.extend(re.findall(r"https?://[^\s]+", text))
        for link in element.findall(".//a:hlinkClick", NS):
            target = relationships.get(link.get(f"{{{NS['r']}}}id"))
            if target:
                found.append(target)
    unique = {}
    for raw in found:
        url = product_url(raw)
        parsed = urllib.parse.urlsplit(url)
        # The visible and embedded links on slide 6 point to the same product
        # with different navigation parameters. Prefer the visible link.
        key = (parsed.hostname, parsed.path)
        unique.setdefault(key, url)
    return list(unique.values())


def retailer(elements, urls):
    if urls:
        host = urllib.parse.urlsplit(urls[0]).hostname
        return STORES.get(host, host.removeprefix("www."))
    text = "\n".join(p for element in elements for p in paragraphs(element))
    match = re.search(r"website\s*:\s*([^\n]+)", text, re.I)
    value = match[1].strip() if match else "Retailer not specified"
    return {"amazon": "Amazon", "ugg": "UGG", "nike": "Nike",
            "tory burch": "Tory Burch"}.get(value.lower(), value)


def extract_products():
    groups = {"clothes": [], "items": [], "golf": []}
    with zipfile.ZipFile(SOURCE) as archive:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        if len(presentation.findall("p:sldIdLst/p:sldId", NS)) != 35:
            raise ValueError("Source slide count changed; review product group boundaries.")
        for slide in [*range(4, 16), *range(17, 27), *range(28, 35)]:
            root = ET.fromstring(archive.read(f"ppt/slides/slide{slide}.xml"))
            rel_root = ET.fromstring(archive.read(f"ppt/slides/_rels/slide{slide}.xml.rels"))
            rels = {r.get("Id"): r.get("Target") for r in rel_root
                    if r.get("TargetMode") == "External"}
            group = "clothes" if slide <= 15 else "items" if slide <= 26 else "golf"
            if slide == 15:
                rows = [row.findall("a:tc", NS) for row in root.findall(".//a:tbl/a:tr", NS)]
                if not rows or len(rows[0]) != 4:
                    raise ValueError("Jewelry table changed; review the item-to-link mapping.")
                records = [([rows[0][column], rows[2][column]],
                            " ".join(p for p in paragraphs(rows[0][column]) if p))
                           for column in range(4)]
            elif slide == 28:
                shapes = [shape for shape in root.findall(".//p:sp", NS)
                          if shape.find("p:nvSpPr/p:cNvPr", NS).get("name")
                          in ("Content Placeholder 3", "Content Placeholder 4")]
                records = [([shape], product_name(shape)) for shape in shapes]
            else:
                title_shape = "TextBox 13" if slide == 9 else "Title"
                title = next(shape for shape in root.findall(".//p:sp", NS)
                             if shape.find("p:nvSpPr/p:cNvPr", NS).get("name", "").startswith(title_shape))
                name = product_name(title)
                records = [([root], name)]
            for elements, name in records:
                if not name:
                    raise ValueError(f"Missing product name on slide {slide}.")
                urls = links(elements, rels)
                groups[group].append({"name": name, "slide": slide,
                                      "size": product_size(elements),
                                      "retailer": retailer(elements, urls), "urls": urls})
    return groups


def render(groups):
    sections = []
    for key, label in (("clothes", "Clothes"), ("items", "Items"), ("golf", "Golf")):
        gifts = groups[key]
        rows = []
        for gift in gifts:
            e = html.escape
            if gift["urls"]:
                destination = "<br>".join(
                    f'<a href="{e(url, quote=True)}" target="_blank" rel="noopener noreferrer">'
                    f'Shop at {e(STORES.get(urllib.parse.urlsplit(url).hostname, gift["retailer"]))}'
                    ' <span aria-hidden="true">↗</span></a>'
                    for url in gift["urls"])
            else:
                destination = (f'{e(gift["retailer"])}'
                               '<span class="missing-link">Product link not provided</span>')
            size_value = WEB_SIZE_OVERRIDES.get((gift["slide"], gift["name"]), gift["size"])
            size = (f'<p class="gift-size"><strong>Size:</strong> {e(display_size(size_value))}</p>'
                    if size_value else "")
            rows.append(
                f'            <li class="gift" data-slide="{gift["slide"]}">\n'
                f'              <div><h4>{e(gift["name"])}</h4>{size}'
                f'<p class="slide-number">Slide {gift["slide"]}</p></div>\n'
                f'              <div class="retailer">{destination}</div>\n'
                '            </li>')
        sections.append(
            f'        <section class="gift-section" id="{key}" aria-labelledby="{key}-heading">\n'
            f'          <div class="category-heading"><h3 id="{key}-heading">{label}</h3>'
            f'<span>{len(gifts)} gift ideas</span></div>\n'
            '          <ul class="gift-list">\n' + "\n".join(rows) +
            '\n          </ul>\n        </section>')
    return "\n".join(sections)


if __name__ == "__main__":
    groups = extract_products()
    source_html = PAGE.read_text(encoding="utf-8")
    output, count = re.subn(
        r"(?<=<!-- BEGIN GENERATED GIFTS -->).*?(?=        <!-- END GENERATED GIFTS -->)",
        lambda _: "\n" + render(groups) + "\n", source_html, flags=re.S)
    if count != 1:
        raise ValueError("Expected exactly one generated gift-list block.")
    PAGE.write_text(output, encoding="utf-8", newline="\n")
    build = ROOT / ".wishlist-build"
    build.mkdir(exist_ok=True)
    (build / "products.json").write_text(json.dumps(groups, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(gifts) for gifts in groups.values())
    linked = sum(bool(gift["urls"]) for gifts in groups.values() for gift in gifts)
    print(f"Generated {total} gift entries; {linked} have source product links.")
