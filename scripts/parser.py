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
        "--dir": "Input directory",
        "--outdir": "Output directory",
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

    def check_values(self):
        arguments = self.parse_args()
        if hasattr(arguments, 'dir'):
            if not os.path.exists(arguments.dir):
                self.error(arguments.dir + " doesn't exists")

        if hasattr(arguments, 'outdir'):
            if not os.path.exists(arguments.outdir):
                self.error(arguments.outdir + " doesn't exists")

        if hasattr(arguments, 'soffice'):
            if not os.path.exists(arguments.soffice):
                self.error(arguments.soffice + " doesn't exists")
            if not arguments.soffice.endswith('/soffice'):
                self.error(arguments.soffice + " should end with '/soffice'")

        if hasattr(arguments, 'type'):
            if arguments.type not in types:
                self.error(arguments.type + " is an invalid type")

        if hasattr(arguments, 'component'):
            if arguments.component not in  components:
                self.error(arguments.component + " is an invalid component")

        return arguments

# vim:set shiftwidth=4 softtabstop=4 expandtab:
