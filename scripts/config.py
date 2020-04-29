config = {
    "ooxml":
    {
        "writer": {
            "import": ["doc", "docx", "rtf"],
            "export": ["doc", "docx", "rtf", "odt"]
        },
        "impress": {
            "import": ["ppt", "pptx"],
            "export": ["ppt", "pptx", "odp"]
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
