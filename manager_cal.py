# Entry-point voor de beheeromgeving van een tweede vestiging.
#
# Zie viewer_cal.py voor de uitleg: dit geeft een unieke main file path zodat
# Streamlit Community Cloud een aparte app kan aanmaken die dezelfde gedeelde
# code draait. De getoonde vestiging/data volgt uit de secrets ([locatie] en
# [supabase]).

import runpy

runpy.run_path("manager_app.py", run_name="__main__")
