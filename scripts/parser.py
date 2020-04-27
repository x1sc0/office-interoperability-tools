# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
# Version: MPL 1.1 / GPLv3+ / LGPLv3+
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License or as specified alternatively below. You may obtain a copy of
# the License at https://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.

import argparse
import os
import sys

components = ["writer", "calc", "impress"]
types = ["odf", "ooxml"]

arguments_descriptions = {
        "--soffice": "soffice instance to connect to",
        "--reference": "Reference directory",
        "--indir": "Input directory",
        "--outdir": "Output directory",
        "--wineprefix": "Path to wineprefix",
        "--extension": "Extension of files to be converted to PDF",
        "--type": "Mimetype to be used. Options: " + \
                    " ".join("[" + x + "]" for x in types),
        "--component": "Component to be used. Options: " + \
                    " ".join("[" + x + "]" for x in components),
}

class CommonParser(argparse.ArgumentParser):

    def error(self, message):
        sys.stderr.write('error: %s\n' % message)
        self.print_help()
        sys.exit(2)

    def add_arguments(self, args):
        for arg in args:
            self.add_argument(arg, required=True,
                help=arguments_descriptions[arg])

    def add_optional_arguments(self, args):
        for arg in args:
            self.add_argument(arg, required=False,
                help=arguments_descriptions[arg])

    def check_path(self, path):
        if not os.path.exists(path):
            self.error(path + " doesn't exists")

        if not os.path.isabs(path):
            self.error(path + " is not an absolute path")

    def check_values(self):
        arguments = self.parse_args()
        if hasattr(arguments, 'indir'):
            self.check_path(arguments.indir)

        if hasattr(arguments, 'reference'):
            self.check_path(arguments.reference)

        if hasattr(arguments, 'outdir'):
            self.check_path(arguments.outdir)

        if hasattr(arguments, 'wineprefix') and arguments.wineprefix:
            self.check_path(arguments.wineprefix)

        if hasattr(arguments, 'soffice'):
            self.check_path(arguments.soffice)

            if not arguments.soffice.endswith('/soffice'):
                self.error(arguments.soffice + " should end with '/soffice'")

        if hasattr(arguments, 'type'):
            if arguments.type not in types:
                self.error(arguments.type + " is an invalid type")

            if arguments.type == "ooxml" and hasattr(arguments, 'wineprefix') and not arguments.wineprefix:
                self.error("wineprefix argument is needed when using 'ooxml' type")

        if hasattr(arguments, 'component'):
            if arguments.component not in  components:
                self.error(arguments.component + " is an invalid component")

        return arguments

# vim:set shiftwidth=4 softtabstop=4 expandtab:
