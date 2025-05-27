from mini_book import MiniBook

book = MiniBook(
    page_content=[
        "Page 1",
        "Page 2",
        "Page 3",
        "Page 4",
        "Page 5",
        "Page 6",
        "Page 7",
        "Page 8",
    ]
)

book.save("mini_book.docx")
