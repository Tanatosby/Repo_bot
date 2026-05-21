"""
Script de debug para SSASECQ.
Abre Banner, navega a SSASECQ, escanea el DOM y prueba llenar un curso de ejemplo.
Uso: python ssasecq_debug.py
"""

import time
from playwright.sync_api import sync_playwright

BANNER_URL    = "https://banner.udep.edu.pe/applicationNavigator/seamless"
PERIODO       = "202610"
TEST_MATERIA  = "ANTR"
TEST_NUMERO   = "01002"

# =============================================================================
# JS helpers
# =============================================================================

JS_SCAN = """
() => {
    const result = { inputs: [], selects: [], buttons: [] };

    for (const inp of document.querySelectorAll('input:not([type="hidden"])')) {
        let label = '';
        if (inp.id) {
            const lbl = document.querySelector('label[for="' + inp.id + '"]');
            if (lbl) label = lbl.textContent.trim();
        }
        if (!label) {
            const parent = inp.closest('td, li, div, span');
            if (parent) {
                const lEl = parent.querySelector('label, span.ng-binding');
                if (lEl && lEl !== inp) label = lEl.textContent.trim().slice(0, 60);
            }
        }
        result.inputs.push({
            id:          inp.id || '',
            name:        inp.name || '',
            type:        inp.type || 'text',
            label:       label.slice(0, 60),
            placeholder: inp.placeholder || '',
            value:       inp.value || '',
            readOnly:    inp.readOnly,
            cls:         inp.className.slice(0, 60)
        });
    }

    for (const sel of document.querySelectorAll('select')) {
        let label = '';
        if (sel.id) {
            const lbl = document.querySelector('label[for="' + sel.id + '"]');
            if (lbl) label = lbl.textContent.trim();
        }
        const options = [...sel.options].map(o => ({ value: o.value, text: o.text.trim() }));
        result.selects.push({
            id:           sel.id || '',
            name:         sel.name || '',
            label:        label,
            currentValue: sel.value,
            cls:          sel.className.slice(0, 60),
            options:      options
        });
    }

    for (const btn of document.querySelectorAll('button, input[type="button"], input[type="submit"]')) {
        result.buttons.push({
            tag:   btn.tagName,
            text:  btn.textContent.trim().slice(0, 40),
            value: btn.value || '',
            id:    btn.id || '',
            cls:   btn.className.slice(0, 60)
        });
    }

    return result;
}
"""

JS_AGREGAR = """
(fieldValue) => {
    const sel = [...document.querySelectorAll('select')].find(
        s => s.options.length > 0 && s.options[0].value === ''
    );
    if (!sel) return 'no_select';
    console.log('Agregar select encontrado, cls=' + sel.className);
    sel.value = fieldValue;
    sel.dispatchEvent(new Event('change', {bubbles: true}));
    return 'ok';
}
"""

JS_FILL_LAST_FILTER = """
(valor) => {
    // Intento 1: inputs con id=framesNN
    const byFrames = [...document.querySelectorAll('input[type="text"]:not([readonly])')].filter(
        i => i.id && /^frames\\d+$/.test(i.id)
    );
    if (byFrames.length) {
        const last = byFrames[byFrames.length - 1];
        last.value = valor;
        last.dispatchEvent(new Event('input',  {bubbles: true}));
        last.dispatchEvent(new Event('change', {bubbles: true}));
        return { method: 'framesNN', id: last.id, name: last.name };
    }

    // Intento 2: cualquier último input editable
    const all = [...document.querySelectorAll('input[type="text"]:not([readonly])')];
    if (all.length) {
        const last = all[all.length - 1];
        last.value = valor;
        last.dispatchEvent(new Event('input',  {bubbles: true}));
        last.dispatchEvent(new Event('change', {bubbles: true}));
        return { method: 'last_editable', id: last.id, name: last.name };
    }
    return null;
}
"""

JS_SET_LAST_OPERATOR = """
(operatorValue) => {
    const opSels = [...document.querySelectorAll('select')].filter(s =>
        s.options[0] && s.options[0].value !== '' &&
        [...s.options].some(o => o.value === '=') &&
        [...s.options].some(o => o.value === '#LIKE')
    );
    if (!opSels.length) return { ok: false, total: 0 };
    const last = opSels[opSels.length - 1];
    last.value = operatorValue;
    last.dispatchEvent(new Event('change', {bubbles: true}));
    return { ok: true, total: opSels.length, id: last.id };
}
"""

JS_CHECK_RESULTS = """
() => {
    // Indicador 1: input de resultado con valor
    const maxEnrl = document.querySelector('input[name="ssbsectMaxEnrl"]');
    if (maxEnrl && maxEnrl.value && maxEnrl.value.trim())
        return { tiene: true, metodo: 'maxEnrl', val: maxEnrl.value };

    // Indicador 2: texto de paginación
    const allText = document.body.innerText || '';
    const m = allText.match(/Registro\\s+\\d+\\s+de\\s+(\\d+)/);
    if (m) return { tiene: parseInt(m[1]) > 0, metodo: 'paginacion', total: parseInt(m[1]) };

    // Indicador 3: cualquier input ssbsect con valor
    const ssbInputs = [...document.querySelectorAll('input[name^="ssbsect"]')]
        .filter(i => i.value && i.value.trim());
    if (ssbInputs.length)
        return { tiene: true, metodo: 'ssbInputs', count: ssbInputs.length,
                 sample: ssbInputs[0].name + '=' + ssbInputs[0].value };

    return { tiene: false, metodo: 'vacio' };
}
"""


def print_scan(data, titulo="SCAN"):
    print(f"\n{'='*60}")
    print(f"  {titulo}")
    print(f"{'='*60}")
    print(f"  Inputs ({len(data['inputs'])}):")
    for inp in data['inputs']:
        ro = " [RO]" if inp.get('readOnly') else ""
        print(f"    [{inp['type']}] id='{inp['id']}' name='{inp['name']}'"
              f" label='{inp['label']}' placeholder='{inp['placeholder']}'"
              f" val='{inp['value']}'{ro}")
        if inp.get('cls'):
            print(f"      cls='{inp['cls']}'")

    print(f"\n  Selects ({len(data['selects'])}):")
    for sel in data['selects']:
        print(f"    id='{sel['id']}' name='{sel['name']}' label='{sel['label']}'"
              f" val='{sel['currentValue']}' cls='{sel['cls']}'")
        for opt in sel['options']:
            marker = " ◄ AGREGAR" if opt['value'] == '' else ""
            print(f"      option value='{opt['value']}' text='{opt['text']}'{marker}")

    print(f"\n  Botones ({len(data['buttons'])}):")
    for btn in data['buttons']:
        print(f"    [{btn['tag']}] id='{btn['id']}' text='{btn['text']}' value='{btn['value']}'")
    print(f"{'='*60}\n")


def get_ssasecq_frame(page):
    """Espera al iframe de SSASECQ."""
    for attempt in range(40):
        for fr in page.frames:
            if fr.url == page.url:
                continue
            try:
                if 'SSASECQ' in fr.url or fr.locator('input[type="text"]').count() >= 1:
                    print(f"  Frame detectado (intento {attempt+1}): {fr.url[:80]}")
                    return fr
            except Exception:
                pass
        time.sleep(0.5)
    print("  ADVERTENCIA: frame no encontrado, usando página principal")
    return None


def run():
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome", headless=False, slow_mo=60)
        except Exception:
            browser = p.chromium.launch(headless=False, slow_mo=60)

        context = browser.new_context()
        page = context.new_page()

        print(f"Abriendo Banner: {BANNER_URL}")
        page.goto(BANNER_URL, wait_until="domcontentloaded", timeout=30000)
        input("\n>>> Inicia sesión en Banner y presiona ENTER <<<\n")
        time.sleep(2)

        # ── Navegar a SSASECQ ──────────────────────────────────────────────
        print("\nNavegando a SSASECQ...")
        page.locator("#sidebarSearchLink").click()
        page.wait_for_selector("#searchMenu", state="visible", timeout=8000)
        time.sleep(0.8)

        search = page.locator("input#search")
        search.wait_for(state="visible", timeout=5000)
        search.click()
        search.fill("SSASECQ")
        time.sleep(2)

        resultado = page.locator("#vsearchResultId li").first
        resultado.wait_for(state="visible", timeout=8000)
        print(f"  Resultado: {resultado.inner_text()[:80]}")
        resultado.click()
        time.sleep(4)
        print(f"  URL tras navegación: {page.url}")

        # ── Detectar frame ────────────────────────────────────────────────
        form = get_ssasecq_frame(page)
        ctx = form if form else page

        # ── SCAN 1: estado inicial ────────────────────────────────────────
        print("\nEsperando que el formulario cargue...")
        time.sleep(3)
        data = ctx.evaluate(JS_SCAN)
        print_scan(data, "SCAN 1 — estado inicial")

        input("\n>>> Revisa el scan inicial y presiona ENTER para continuar <<<\n")

        # ── Intentar llenar Periodo ────────────────────────────────────────
        print(f"\nIntentando llenar Periodo con '{PERIODO}'...")

        # Buscar el input de Periodo: primero por label, luego por id conocido
        filled_periodo = ctx.evaluate(f"""
            () => {{
                // Buscar por label 'Periodo' en los filter inputs (framesNN)
                const filterInputs = [...document.querySelectorAll(
                    'input[type="text"]:not([readonly])'
                )].filter(i => i.id && /^frames\\d+$/.test(i.id));

                if (filterInputs.length > 0) {{
                    const inp = filterInputs[0];
                    inp.value = '{PERIODO}';
                    inp.dispatchEvent(new Event('input',  {{bubbles: true}}));
                    inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                    return {{ ok: true, id: inp.id, name: inp.name }};
                }}

                // Fallback: primer input editable que no sea de búsqueda
                const all = [...document.querySelectorAll('input[type="text"]:not([readonly])')].filter(
                    i => !i.placeholder.includes('Buscar')
                );
                if (all.length) {{
                    const inp = all[0];
                    inp.value = '{PERIODO}';
                    inp.dispatchEvent(new Event('input',  {{bubbles: true}}));
                    inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                    return {{ ok: true, method: 'fallback', id: inp.id, name: inp.name }};
                }}
                return {{ ok: false }};
            }}
        """)
        print(f"  Periodo fill: {filled_periodo}")
        time.sleep(0.5)

        # ── SCAN 2: después de llenar Periodo ─────────────────────────────
        data2 = ctx.evaluate(JS_SCAN)
        print_scan(data2, "SCAN 2 — tras llenar Periodo")

        input("\n>>> Revisa scan post-Periodo y presiona ENTER para agregar Materia <<<\n")

        # ── Agregar campo Materia ─────────────────────────────────────────
        print(f"\nAgregando campo Materia (SSBSECT_SUBJ_CODE)...")
        ok = ctx.evaluate(JS_AGREGAR, 'SSBSECT_SUBJ_CODE')
        print(f"  Agregar Materia: {ok}")
        time.sleep(1.5)

        op = ctx.evaluate(JS_SET_LAST_OPERATOR, '=')
        print(f"  Operador Materia → '=': {op}")
        time.sleep(0.3)

        filled = ctx.evaluate(JS_FILL_LAST_FILTER, TEST_MATERIA)
        print(f"  Fill Materia='{TEST_MATERIA}': {filled}")
        time.sleep(0.5)

        # ── SCAN 3: después de agregar Materia ────────────────────────────
        data3 = ctx.evaluate(JS_SCAN)
        print_scan(data3, "SCAN 3 — tras agregar Materia")

        input("\n>>> Revisa scan post-Materia y presiona ENTER para agregar Curso <<<\n")

        # ── Agregar campo Curso ───────────────────────────────────────────
        print(f"\nAgregando campo Curso (SSBSECT_CRSE_NUMB)...")
        ok = ctx.evaluate(JS_AGREGAR, 'SSBSECT_CRSE_NUMB')
        print(f"  Agregar Curso: {ok}")
        time.sleep(1.5)

        op = ctx.evaluate(JS_SET_LAST_OPERATOR, '=')
        print(f"  Operador Curso → '=': {op}")
        time.sleep(0.3)

        filled = ctx.evaluate(JS_FILL_LAST_FILTER, TEST_NUMERO)
        print(f"  Fill Curso='{TEST_NUMERO}': {filled}")
        time.sleep(0.5)

        # ── SCAN 4: formulario completo antes de ejecutar ─────────────────
        data4 = ctx.evaluate(JS_SCAN)
        print_scan(data4, "SCAN 4 — formulario listo (antes de F8)")

        input("\n>>> Revisa scan final y presiona ENTER para ejecutar F8 <<<\n")

        # ── Ejecutar F8 ───────────────────────────────────────────────────
        print("\nEjecutando F8...")
        page.keyboard.press("F8")
        time.sleep(10)

        # ── SCAN 5: resultados ────────────────────────────────────────────
        data5 = ctx.evaluate(JS_SCAN)
        print_scan(data5, "SCAN 5 — tras F8 (resultados)")

        resultado_check = ctx.evaluate(JS_CHECK_RESULTS)
        print(f"\n  CHECK RESULTADOS: {resultado_check}")

        input("\n>>> Presiona ENTER para cerrar el navegador <<<\n")
        browser.close()


if __name__ == "__main__":
    run()
