from mini_book import MiniBook

book = MiniBook(
    page_content=[
        "<img>images/title.jpg</img>",
        "Wow! Ein Regenbogen! So viele Farben am Himmel. Komm, wir schauen ihn uns genauer an!",
        "<img>images/rainbow.jpg</img>",
        "Wenn die Sonne auf Regentropfen scheint, passiert etwas Magisches! Dann entsteht ein Regenbogen.",
        "<img>images/colors.jpg</img>",
        "Rot, Orange, Gelb, Grün, Blau und Violett - das sind die Farben vom Regenbogen. Kennst du sie alle?",
        "<img>images/animals.jpg</img>",
        "Am Ende vom Regenbogen soll ein Schatz versteckt sein! Vielleicht ein Topf voller bunter Seifenblasen? Pffft... platsch!",
    ]
)
book.save("output/mini_book.docx")

# NOTE: only works on Windows and macOS
# book.export_to_pdf("output/mini_book.pdf")
