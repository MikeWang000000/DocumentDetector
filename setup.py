# -*- coding: GBK -*-

# 使用py2exe编译打包本程序。
# 由于某种诡异的兼容性问题，该程序只能在Windows XP环境下使用py2exe编译打包，或将"bundle_files"设为3。

from distutils.core import setup;
import os;
import sys;
import py2exe;
import CreateIcon;

if len(sys.argv) == 1:
    sys.argv.append("py2exe");

icon_path = os.path.join(os.environ['TEMP'],'dd_icon.ico');
CreateIcon.create(icon_path);

setup(  description = "Document Detector",
        name = "Document Detector",
        options = {"py2exe": {"compressed": 1,
                              "optimize": 2,
                              "ascii": 0,
                              "bundle_files": 1
                             }
                  },
        zipfile = None,
        windows = [{"script": "DocumentDetector.py",
                    "icon_resources": [ (1, icon_path) ]
                  }]  
     );

raw_input("Done.")
