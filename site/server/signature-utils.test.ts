import { describe, expect, it } from "vitest";
import { countLogicalPages, escapeXml } from "../client/src/pages/Home";

describe("signature document helpers", () => {
  it("counts page breaks as logical pages", () => {
    const xml = '<w:document><w:body><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:lastRenderedPageBreak/></w:p><w:p><w:r><w:t>Conteúdo final</w:t></w:r></w:p></w:body></w:document>';
    expect(countLogicalPages(xml)).toBe(3);
  });

  it("ignores a trailing page break with no content after it", () => {
    const xml = '<w:document><w:body><w:p><w:r><w:t>Página 1</w:t><w:br w:type="page"/></w:r></w:p></w:body></w:document>';
    expect(countLogicalPages(xml)).toBe(1);
  });

  it("always exposes at least one page", () => {
    expect(countLogicalPages("<w:document><w:body><w:p/></w:body></w:document>")).toBe(1);
  });

  it("escapes XML attributes safely", () => {
    expect(escapeXml('A & B <draft> "signed"')).toBe("A &amp; B &lt;draft&gt; &quot;signed&quot;");
  });
});
