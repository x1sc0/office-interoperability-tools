config = {
    "ooxml":
    {
        "writer": {
            "import": ["doc", "docx", "rtf"],
            "export": ["pdf", "doc", "docx", "rtf", "odt"]
        },
        "impress": {
            "import": ["ppt", "pptx"],
            "export": ["pdf", "ppt", "pptx", "odp"]
        }
    },
    "odf":
    {
        "writer": {
            "import": ["odt"],
            "export": ["doc", "docx", "rtf", "odt"]
        },
        "impress": {
            "import": ["odp"],
            "export": ["ppt", "pptx", "odp"]
        },
        "calc": {
            "import": ["ods"],
            "export": ["xlsx", "xls", "ods"]
        }
    }
}
