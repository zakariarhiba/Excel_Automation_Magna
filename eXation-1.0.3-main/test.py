import pandas as pd
import numpy as np
import time

tic = time.time()

Export = pd.read_csv(
            ".\csv_files\Export.csv",
            low_memory=False,
            dtype=str,
            encoding="unicode_escape",
        )


Export.to_excel("./out.xlsx", index=False)


toc = time.time()

print(f"done in {toc-tic} seconds")

