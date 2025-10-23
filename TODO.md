To'liq loyiha ishlashi uchun quyidagi 3 ta logical service kerak bo'ladi:

### 1. Parser modulida qilinadigan ishlar:

- [x] ZIP ochuvchi `DocxPackage` loaderni kengaytirib, barcha kerakli XML va media fayllarni cache qilish.
- [x] `rels_parser`da hujjat, media, header/footer, numbering, hyperlink rels xaritalarini to‘liq yig‘ish.
- [x] `styles_parser`da paragraph/character/table style’larni, `baseOn`, `linkedStyle`, `default` atributlarini va `rPr`/`pPr`/`tblPr` ni hisobga olib flatten qilish.
- [ ] `numbering` modulini qo‘shib, `<w:abstractNum>` va `<w:num>` ma’lumotlarini strukturaga yig‘ish.
- [ ] `document_parser`da `<w:body>` oqimini yurib, paragraph, run, text, field, bookmark, table, list, drawing/image, shape, footnote/reference, section break kabi elementlarni modulga mos model obyektlariga aylantirish.
- [ ] Header/footer parserlarini qo‘shib, `sectPr` orqali sahifa bo‘limlariga bog‘lash.
- [ ] Media extractor’da rasmlar, embedded fontlar va boshqa binary assetlarga yo‘l va metadata berish.
- [ ] Parserdan chiqqan elementlarni intermediate modelga yig‘ishda namespace/special charlarini to‘g‘ri dekodlash va whitespace normalizatsiyasini amalga oshirish.
- [ ] Har bir asosiy blok uchun minimal unit testlar yozish va sample DOCX bilan integratsion tekshiruv qo‘shish.