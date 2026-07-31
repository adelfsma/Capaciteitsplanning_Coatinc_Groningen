# Entry-point voor een tweede vestiging (bijv. Coatinc Alblasserdam).
#
# Streamlit Community Cloud identificeert een app op repo + branch + main file.
# Omdat viewer_app.py al door de Groningen-app bezet is, geeft dit bestand een
# UNIEKE main file path, zodat er een aparte app aangemaakt kan worden die
# dezelfde gedeelde code draait. Welke vestiging/data getoond wordt, bepaalt
# volledig de [locatie]- en [supabase]-configuratie in de Streamlit-secrets.
#
# runpy voert viewer_app.py bij elke Streamlit-rerun opnieuw uit (een gewone
# 'import' zou door module-caching maar één keer draaien).

import runpy

runpy.run_path("viewer_app.py", run_name="__main__")
