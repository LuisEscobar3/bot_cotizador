# -*- coding: utf-8 -*-
# ======================================================================
# 🚀 SIMONWEB – FIREFOX – MULTI-TABS – LOCAL (Python + Playwright)
# ======================================================================

import asyncio
import json
import inspect
import unicodedata
from typing import Optional, List

from playwright.async_api import (
    async_playwright,
    Browser,
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
)

# =========================
# CONFIG
# =========================
CONFIG = {
    "login_url": "http://10.1.0.184:3005/SimonWeb/login.html",
    "filtros_url": "http://10.1.0.184:3005/SimonWeb/pages/emision/poliza/paginaFiltros.jsf?modoPantalla=CREAR_COTIZA",

    # Credenciales y compañía
    "usuario": "1070950316",
    "clave": "Ujm17ikl",
    "compania_id": "3",

    # Datos para frmDatosFijos / frmTerceroNatural
    "fecha_inicio_vigencia": "14/11/2025",
    "numero_documento_tercero": "1032497498",

    # Datos del tercero (para alta nueva)
    "primer_nombre": "LUIS",
    "segundo_nombre": "CARLOS",
    "primer_apellido": "ESCOBAR",
    "segundo_apellido": "MUÑOZ",
    "fecha_nacimiento": "01/01/1990",  # dd/mm/yyyy
    "sexo": "M",
    "estado_civil": "1",
    "ciudad_residencia": "BOGOTÁ",

    # Contacto
    "celular_tercero": "3138926061",
    "telefono_residencia_tercero": "1234567",
    "direccion_residencia_tercero": "Calle 123 #45-67",
    "email_tercero": "correo@ejemplo.com",

    # frmDatosFijos
    "tipo_envio": "PE",
    "clave_lider": "31194",
    "age_pactado": "6",

    # Localidad
    "localidad": "BOGOTA",

    # Datos en frmRiesgosPoliza
    "placa": "NHM747",

    # Datos manuales de vehículo
    "auto_marca_manual": "9201276 - VOLKSWAGEN VOYAGE COMFORTLINE TP 1600CC 16V 2AB ABS R15",

    # 👉 USO: TEXTO COMPLETO del option
    "auto_uso_value": "31 - PARTICULAR FAMILIAR",

    # Modelo por texto EXACTO del option:
    "auto_modelo_value": "2021 - MODELO 2021",

    "auto_color": "GRIS",
    "auto_motor": "MOTOR123456",
    "auto_chasis": "CHASIS123456",
    "auto_kilometraje": "0",
    "auto_suma_asegurada": "40000000",
    "auto_fecha_matricula": "01/01/2021",
    "auto_cobertura_value": "43 - OPCIONES - Riesgos Patrimoniales , Daño Total , Daños Parciales , Terremoto , Hurto Parcial Y Total",

    # Ventana / tiempos
    "viewport": {"width": 1900, "height": 1000},
    "timeout_general": 30000,
    "headless": False,
    "slow_mo": 0,

    "keep_open": True,
    "segundo_select_css": "#frmPaginaFiltros\\:programa",
}


# =========================
# PRINT con línea y función
# =========================
def pr(msg: str, tipo: str = "INFO"):
    f = inspect.currentframe().f_back
    print(f"[{tipo}] (L{f.f_lineno}, {f.f_code.co_name}) {msg}", flush=True)


def strip_accents(text: str) -> str:
    if not text:
        return text
    nfkd = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in nfkd if unicodedata.category(ch) != "Mn")


async def safe_fill_and_tab(page: Page, selector: str, value: str, desc: str) -> bool:
    try:
        await page.wait_for_selector(selector, timeout=5000)
        await page.fill(selector, value)
        pr(f"{desc}: valor '{value}' escrito en {selector}")
        await page.press(selector, "Tab")
        pr(f"{desc}: Tab presionado en {selector}")
        return True
    except Exception as e:
        pr(f"{desc}: no se pudo llenar o hacer Tab en {selector} -> {e}", "WARN")
        return False


# =========================
# Utilidades
# =========================
async def wait(ms: int):
    await asyncio.sleep(ms / 1000)


async def wait_for_loader(p: Page, timeout: int = 8000) -> bool:
    pr("Esperando loader…")
    js = """
    () => {
      if (document.querySelector(".rf-st-start") || document.querySelector(".rf-st-rich")) return true;
      const ids = ["#loadingImage", "#ajaxStatusPanel", "#statusPanel", "#statusPanel_container"];
      return ids.some(sel => {
        const el = document.querySelector(sel);
        return el && el.offsetParent !== null;
      });
    }"""
    step = 200
    elapsed = 0
    while elapsed < timeout:
        try:
            active = await p.evaluate(js)
        except Exception:
            active = True
        if not active:
            return True
        await asyncio.sleep(step / 1000)
        elapsed += step
    pr("Loader no desapareció.", "WARN")
    return False


async def safe_click(page: Page, selectors: List[str], desc: str) -> bool:
    for sel in selectors:
        try:
            await page.wait_for_selector(sel, timeout=4000)
            loc = page.locator(sel)
            await loc.scroll_into_view_if_needed()
            await loc.click()
            pr(f"Click en {desc} ({sel})")
            return True
        except Exception:
            continue
    pr(f"No se pudo hacer click en {desc}.", "WARN")
    return False


async def safe_fill(page: Page, selectors: List[str], value: str, desc: str) -> bool:
    for sel in selectors:
        try:
            await page.fill(sel, value, timeout=1500)
            pr(f"Campo {desc} rellenado en {sel}")
            return True
        except Exception:
            continue
    pr(f"No se encontró campo para {desc}.")
    return False


# ============ bloquear cierre de pestañas ============
INIT_SCRIPT_BLOCK_CLOSE_HARD = r"""
(() => {
  const noop = () => {};
  try {
    const targets = [window, self];
    try { if (window.top) targets.push(window.top); } catch(e){}
    try { if (window.parent) targets.push(window.parent); } catch(e){}
    try { if (window.opener) targets.push(window.opener); } catch(e){}
    for (const t of targets) { try { t.close = noop; } catch(e){} }
    const _open = window.open;
    window.open = function(url, name, specs) {
      const w = _open ? _open.call(window, url, name, specs) : null;
      try { if (w) { w.close = noop; } } catch(e){}
      return w;
    };
    window.addEventListener('beforeunload', (e) => { e.preventDefault(); e.returnValue=''; return ''; });
  } catch(e){}
})();
"""


def attach_reopen_on_close(context: BrowserContext, label: str, url_to_restore: Optional[str] = None):
    def _wrap(page: Page):
        async def _reopen():
            try:
                last_url = url_to_restore or page.url
            except Exception:
                last_url = url_to_restore
            print(f"[WARN] (reopen) '{label}' cerrada. Reabriendo {last_url or '(en blanco)'}…", flush=True)
            newp = await context.new_page()
            await newp.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
            if last_url:
                try:
                    await newp.goto(last_url, wait_until="load", timeout=15000)
                    print(f"[INFO] (reopen) '{label}' reabierta en {last_url}", flush=True)
                except Exception as e:
                    print(f"[ERROR] (reopen) Falló reabrir: {e}", flush=True)

        page.on("close", lambda: asyncio.ensure_future(_reopen()))

    return _wrap


async def open_url_resilient(context: BrowserContext, url: str, label: str, max_tries: int = 4) -> Page:
    for i in range(1, max_tries + 1):
        print(f"[INFO] (open_url_resilient) {label} intento {i}/{max_tries}", flush=True)
        p = await context.new_page()
        await p.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
        attach_reopen_on_close(context, label, url)(p)
        try:
            await p.goto(url, wait_until="load", timeout=15000)
            await asyncio.sleep(0.4)
            if p.is_closed():
                print("[WARN] (open_url_resilient) Pestaña se cerró tras cargar, reintentando…", flush=True)
                continue
            return p
        except Exception as e:
            print(f"[WARN] (open_url_resilient) goto falló: {e}", flush=True)
            try:
                if not p.is_closed():
                    await p.close()
            except:
                pass
            await asyncio.sleep(0.3)
    last = await context.new_page()
    await last.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
    attach_reopen_on_close(context, label, url)(last)
    return last


async def handle_acceso_denegado(context: BrowserContext, current_tab: Page, url: str) -> Page:
    try:
        html = await current_tab.content()
    except Exception:
        html = ""
    if ("Acceso Denegado" in html) or ("denegado" in html.lower()):
        print("[WARN] Acceso Denegado detectado.", flush=True)
        await asyncio.sleep(1.0)
        for p in context.pages:
            u = (p.url or "").lower()
            if p is not current_tab and (u == "about:blank" or u.startswith("moz-extension://") or "deneg" in u):
                print("[INFO] Usando pestaña auto-abierta…", flush=True)
                await p.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
                attach_reopen_on_close(context, "auto", url)(p)
                await p.bring_to_front()
                try:
                    await p.goto(url, wait_until="load", timeout=15000)
                    await asyncio.sleep(0.3)
                    if not p.is_closed():
                        return p
                except Exception as e:
                    print(f"[WARN] Falló auto-abierta: {e}", flush=True)
        print("[INFO] Creando nueva pestaña para reintentar…", flush=True)
        return await open_url_resilient(context, url, "retry")
    return current_tab


# =========================
# Selects RichFaces/JSF
# =========================
async def select_richfaces_value(page: Page, css_select: str, value: str, wait_ajax_ms: int = 1200) -> bool:
    """
    Selecciona un <select> por value. Además, si no coincide exacto,
    intenta buscar option cuyo value empiece por el texto dado.
    """
    try:
        await page.wait_for_selector(css_select, state="visible", timeout=7000)
    except Exception:
        print(f"[WARN] No encontré el select {css_select}", flush=True)
        return False

    # 1) Intento directo
    try:
        await page.select_option(css_select, value)
    except Exception:
        pass

    # 2) Verificar qué quedó
    try:
        curr = await page.eval_on_selector(css_select, "el => el.value")
    except Exception:
        curr = None

    if curr != value:
        # 3) Fallback JS
        result = await page.evaluate(
            """(args) => {
                const { sel, val } = args;
                const s = document.querySelector(sel);
                if (!s) return null;
                const norm = (v) => (v || '').trim();
                const target = norm(val);
                const opts = [...s.options];

                let opt = opts.find(o => norm(o.value) === target);
                if (!opt) {
                    opt = opts.find(o => norm(o.value).startsWith(target));
                }
                if (!opt) return null;

                s.value = opt.value;
                s.dispatchEvent(new Event('change', { bubbles: true }));
                s.dispatchEvent(new Event('input',  { bubbles: true }));
                s.blur && s.blur();

                const selected = s.options[s.selectedIndex] || opt;
                return {
                    value: selected.value || '',
                    text: selected.text || ''
                };
            }""",
            {"sel": css_select, "val": value}
        )

        if result:
            curr = result["value"]
            print(f"[INFO] {css_select} (JS fallback) -> value='{result['value']}' text='{result['text']}'",
                  flush=True)
        else:
            print(f"[WARN] JS fallback no encontró opción para value='{value}' en {css_select}", flush=True)
    else:
        print(f"[INFO] {css_select} -> set={curr} target={value}", flush=True)

    await asyncio.sleep(wait_ajax_ms / 1000)
    return curr is not None and curr != ""


async def select_any_with_value(page: Page, value: str, wait_ajax_ms: int = 1200) -> bool:
    sel = await page.evaluate(
        """(args) => {
            const { val } = args;
            const sels = [...document.querySelectorAll('select')].filter(s => {
              const cs = getComputedStyle(s);
              return s.offsetParent !== null && cs.display !== 'none' && cs.visibility !== 'hidden';
            });
            for (const s of sels) {
              if ([...s.options].some(o => o.value === val)) {
                if (s.id) return `#${CSS.escape(s.id)}`;
                if (s.name) return `select[name="${s.name}"]`;
                const idx = sels.indexOf(s);
                return `select:nth-of-type(${idx+1})`;
              }
            }
            return null;
        }""",
        {"val": value}
    )
    if not sel:
        print(f"[WARN] No hallé ningún <select> con option[value='{value}']", flush=True)
        return False
    return await select_richfaces_value(page, sel, value, wait_ajax_ms)


async def select_option_by_text(
    page: Page,
    css_select: str,
    text: str,
    desc: str,
    wait_ajax_ms: int = 1200,
    max_wait_ms: int = 20000,
) -> bool:
    """
    Selecciona <option> por TEXTO visible (innerText).
    Espera a que carguen opciones (AJAX) → mínimo > 1 opción.
    """
    try:
        await page.wait_for_selector(css_select, state="visible", timeout=7000)
    except Exception:
        pr(f"[{desc}] No encontré el select {css_select}", "WARN")
        return False

    elapsed = 0
    step = 500
    first_dump = True

    while elapsed <= max_wait_ms:
        result = await page.evaluate(
            """
            (args) => {
                const { sel, text } = args;
                const s = document.querySelector(sel);
                if (!s) return { found: false, options: [] };

                const norm = v => (v || '')
                    .replace(/\\u00a0/g, ' ')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toUpperCase();

                const target = norm(text);
                const opts = Array.from(s.options).map(o => ({
                    text: o.text,
                    value: o.value
                }));

                // Si solo está "--Seleccione--", aún no cargan las opciones reales
                if (opts.length <= 1) {
                    return { found: false, options: opts };
                }

                let opt = opts.find(o => norm(o.text) === target);
                if (!opt) opt = opts.find(o => norm(o.text).includes(target));
                if (!opt) return { found: false, options: opts };

                const realOpt = Array.from(s.options).find(
                    o => o.text === opt.text && o.value === opt.value
                );
                if (!realOpt) return { found: false, options: opts };

                s.value = realOpt.value;
                s.dispatchEvent(new Event('change', { bubbles: true }));
                s.dispatchEvent(new Event('input',  { bubbles: true }));
                if (typeof s.blur === 'function') s.blur();

                return {
                    found: true,
                    selected: {
                        value: realOpt.value || '',
                        text: realOpt.text || ''
                    },
                    options: opts
                };
            }
            """,
            {"sel": css_select, "text": text}
        )

        if result and result.get("found"):
            sel_info = result.get("selected", {})
            pr(
                f"[{desc}] seleccionado -> value='{sel_info.get('value', '')}' "
                f"text='{sel_info.get('text', '')}' en {css_select}"
            )
            await asyncio.sleep(wait_ajax_ms / 1000)
            return True

        if first_dump and result:
            first_dump = False
            opts = result.get("options", [])
            pr(
                f"[{desc}] opciones visibles en {css_select}: "
                + ", ".join([f"'{o['text']}'" for o in opts])
            )

        await asyncio.sleep(step / 1000)
        elapsed += step

    pr(f"[{desc}] No se encontró opción con texto '{text}' en {css_select}", "WARN")
    return False


# =========================
# MAIN
# =========================
async def main():
    browser: Browser = None
    context: BrowserContext = None
    main_page: Page = None
    popup: Optional[Page] = None
    filtros_tab: Optional[Page] = None

    success = False
    url_final = None
    err = None

    try:
        pr("Lanzando Firefox local…")
        async with async_playwright() as pw:
            browser = await pw.firefox.launch(
                headless=CONFIG["headless"],
                slow_mo=CONFIG["slow_mo"],
            )
            context = await browser.new_context()
            await context.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
            context.set_default_timeout(9000)
            context.set_default_navigation_timeout(15000)

            # Ventana principal
            main_page = await context.new_page()
            await main_page.set_viewport_size(CONFIG["viewport"])
            attach_reopen_on_close(context, "main")(main_page)

            # Login
            pr("Navegando a login…")
            await main_page.goto(CONFIG["login_url"], wait_until="load", timeout=CONFIG["timeout_general"])

            user_filled = await safe_fill(main_page, ["#Num_Documento", "input[name='Ecom_User_ID']"],
                                          CONFIG["usuario"], "Usuario")
            pass_filled = await safe_fill(main_page, ["input[type='password']", "input[name='Ecom_Password']"],
                                          CONFIG["clave"], "Contraseña")

            if user_filled or pass_filled:
                _ = await safe_click(main_page, ["input#submit", "button#submit"], "Login")
                try:
                    popup = await context.wait_for_event("page", timeout=12000)
                    await popup.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
                    attach_reopen_on_close(context, "popup")(popup)
                    await popup.wait_for_load_state("load")
                except PlaywrightTimeoutError:
                    pr("No se abrió popup; sigo en principal.")
                    popup = main_page
            else:
                pr("Intento disparar popup de login…")
                evt = context.wait_for_event("page", timeout=12000)
                await safe_click(main_page, ["input#submit"], "Login (abre popup)")
                try:
                    popup = await evt
                    await popup.add_init_script(INIT_SCRIPT_BLOCK_CLOSE_HARD)
                    attach_reopen_on_close(context, "popup")(popup)
                    await popup.wait_for_load_state("load")
                except PlaywrightTimeoutError:
                    pr("No apareció popup; sigo en principal.", "WARN")
                    popup = main_page

            # Popup / selección de compañía
            if popup:
                await safe_fill(popup, ["input[name='Ecom_User_ID']"], CONFIG["usuario"], "Usuario (popup)")
                await safe_fill(popup, ["input[name='Ecom_Password']"], CONFIG["clave"], "Contraseña (popup)")
                await safe_click(popup, ["#j_idt138", "button:has-text('OK')", "input[value='OK']"], "OK")
                await wait(600)
                await wait_for_loader(popup, timeout=7000)

                try:
                    await popup.wait_for_selector("#companyForm\\:companyList", timeout=7000)
                    await popup.select_option("#companyForm\\:companyList", CONFIG["compania_id"])
                    await wait_for_loader(popup, timeout=7007)
                except Exception:
                    pr("No hay selector de compañía (o ya está seteado).")

                await safe_click(popup,
                                 ["#companyForm\\:j_idt129", "button:has-text('Continuar')",
                                  "input[value='Continuar']"],
                                 "Continuar")
                await wait_for_loader(popup, timeout=7000)

            # Abrir filtros
            print("[INFO] Abriendo paginaFiltros.jsf en pestaña nueva…", flush=True)
            filtros_tab = await open_url_resilient(context, CONFIG["filtros_url"], "filtros", max_tries=4)
            filtros_tab = await handle_acceso_denegado(context, filtros_tab, CONFIG["filtros_url"])

            if filtros_tab.is_closed():
                print("[ERROR] La pestaña de filtros está cerrada.", flush=True)
            else:
                await filtros_tab.bring_to_front()
                print("[INFO] Filtros abiertos / pestaña viva.", flush=True)

                # 1) producto 250
                PRODUCTO_SEL = "#frmPaginaFiltros\\:producto"
                print("[INFO] Seleccionando 250 en frmPaginaFiltros:producto…", flush=True)
                ok_250 = await select_richfaces_value(filtros_tab, PRODUCTO_SEL, "250", wait_ajax_ms=1400)
                if not ok_250:
                    print("[WARN] No se pudo seleccionar 250 en frmPaginaFiltros:producto", flush=True)
                else:
                    print("[OK] Seleccionado 250 en frmPaginaFiltros:producto", flush=True)

                # 2) programa 251
                print("[INFO] Intentando seleccionar 251 (programa)…", flush=True)
                ok_251 = False
                try:
                    ok_251 = await select_richfaces_value(
                        filtros_tab, CONFIG["segundo_select_css"], "251", wait_ajax_ms=1200
                    )
                except Exception:
                    ok_251 = False

                if not ok_251:
                    print("[INFO] Reintentando 251 por búsqueda dinámica…", flush=True)
                    ok_251 = await select_any_with_value(filtros_tab, "251", wait_ajax_ms=1200)

                if ok_251:
                    print("[OK] Seleccionado 251.", flush=True)
                else:
                    print("[WARN] No se pudo seleccionar 251.", flush=True)

                # 3) botón filtros
                print("[INFO] Haciendo click en frmPaginaFiltros:j_idt322…", flush=True)
                click_ok = await safe_click(
                    filtros_tab,
                    [
                        "#frmPaginaFiltros\\:j_idt322",
                        "xpath=//*[@id='frmPaginaFiltros:j_idt322']",
                    ],
                    "Botón frmPaginaFiltros:j_idt322"
                )
                if click_ok:
                    await wait_for_loader(filtros_tab, timeout=8000)
                else:
                    print("[WARN] No se pudo hacer click en frmPaginaFiltros:j_idt322.", flush=True)

                # Fecha + doc
                FECHA_SEL = "xpath=//*[@id='frmDatosFijos:fechaInicioVigenciaInputDate']"
                await safe_fill_and_tab(filtros_tab, FECHA_SEL,
                                        CONFIG["fecha_inicio_vigencia"], "Fecha Inicio Vigencia")

                NUM_DOC_SEL = "xpath=//*[@id='frmDatosFijos:componenteTercero:numeroDocumento']"
                await safe_fill_and_tab(filtros_tab, NUM_DOC_SEL,
                                        CONFIG["numero_documento_tercero"], "Número Documento Tercero")

                # mensajeModificacion
                msg_sel = "xpath=//*[@id='frmTerceroNatural:mensajeModificacion']"
                msg_modif = False
                msg_info_actualizada = False

                try:
                    await filtros_tab.wait_for_selector(msg_sel, timeout=8000)
                    el = filtros_tab.locator(msg_sel)
                    msg_text = (await el.inner_text()).strip()
                    print(f"[INFO] mensajeModificacion encontrado: '{msg_text}'", flush=True)
                    lower_msg = msg_text.lower()
                    if "información actualizada" in lower_msg or "informacion actualizada" in lower_msg:
                        msg_modif = True
                        msg_info_actualizada = True
                    elif any(w in lower_msg for w in ["modific", "actualiz", "datos", "tercero"]):
                        msg_modif = True
                except Exception:
                    print("[INFO] No apareció mensajeModificacion → TERCERO NUEVO.", flush=True)

                # Tercero nuevo / existente
                if not msg_modif:
                    # NUEVO
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:primerNombre']"],
                        CONFIG["primer_nombre"],
                        "Primer Nombre"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:segundoNombre']"],
                        CONFIG["segundo_nombre"],
                        "Segundo Nombre"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:primerApellido']"],
                        CONFIG["primer_apellido"],
                        "Primer Apellido"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:segundoApellido']"],
                        CONFIG["segundo_apellido"],
                        "Segundo Apellido"
                    )

                    await safe_fill_and_tab(
                        filtros_tab,
                        "xpath=//*[@id='frmTerceroNatural:fechaNacimientoInputDate']",
                        CONFIG["fecha_nacimiento"],
                        "Fecha nacimiento"
                    )

                    sexo_val = CONFIG.get("sexo", "M").upper()
                    sexo_id = (
                        "xpath=//*[@id='frmTerceroNatural:sexo:0']"
                        if sexo_val == "M"
                        else "xpath=//*[@id='frmTerceroNatural:sexo:1']"
                    )
                    await safe_click(filtros_tab, [sexo_id], f"Sexo {sexo_val}")

                    ESTADO_SEL = "xpath=//*[@id='frmTerceroNatural:estadoCivil']"
                    try:
                        await filtros_tab.select_option(ESTADO_SEL, CONFIG["estado_civil"])
                    except Exception as e:
                        print(f"[WARN] No se pudo seleccionar estado civil → {e}", flush=True)

                    ciudad_plain = strip_accents(CONFIG["ciudad_residencia"]).upper()
                    CIUDAD_SEL = "xpath=//*[@id='frmTerceroNatural:ciudadResidenciaInput']"
                    try:
                        await filtros_tab.wait_for_selector(CIUDAD_SEL, timeout=8000)
                        ciudad_input = filtros_tab.locator(CIUDAD_SEL)
                        await ciudad_input.click()
                        await ciudad_input.fill("")
                        await ciudad_input.type(ciudad_plain, delay=150)
                        await filtros_tab.press(CIUDAD_SEL, "Enter")
                    except Exception as e:
                        print(f"[WARN] No se pudo escribir ciudad residencia (nuevo) → {e}", flush=True)

                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:celular']"],
                        CONFIG["celular_tercero"],
                        "Celular tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:telefonoResidencia']"],
                        CONFIG["telefono_residencia_tercero"],
                        "Teléfono residencia tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:direccionResidencia']"],
                        CONFIG["direccion_residencia_tercero"],
                        "Dirección residencia tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:email']"],
                        CONFIG["email_tercero"],
                        "Email tercero"
                    )

                    do_guardar = True
                    do_cancelar = False

                else:
                    # EXISTENTE
                    ciudad_plain = strip_accents(CONFIG["ciudad_residencia"]).upper()
                    CIUDAD_SEL = "xpath=//*[@id='frmTerceroNatural:ciudadResidenciaInput']"
                    try:
                        await filtros_tab.wait_for_selector(CIUDAD_SEL, timeout=8000)
                        ciudad_input = filtros_tab.locator(CIUDAD_SEL)
                        await ciudad_input.click()
                        await ciudad_input.fill("")
                        await ciudad_input.type(ciudad_plain, delay=150)
                        await filtros_tab.press(CIUDAD_SEL, "Enter")
                    except Exception as e:
                        print(f"[WARN] No se pudo escribir ciudad residencia (existente) → {e}", flush=True)

                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:celular']"],
                        CONFIG["celular_tercero"],
                        "Celular tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:telefonoResidencia']"],
                        CONFIG["telefono_residencia_tercero"],
                        "Teléfono residencia tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:direccionResidencia']"],
                        CONFIG["direccion_residencia_tercero"],
                        "Dirección residencia tercero"
                    )
                    await safe_fill(
                        filtros_tab,
                        ["xpath=//*[@id='frmTerceroNatural:email']"],
                        CONFIG["email_tercero"],
                        "Email tercero"
                    )

                    if msg_info_actualizada:
                        do_guardar = False
                        do_cancelar = True
                    else:
                        do_guardar = True
                        do_cancelar = False

                # Guardar / cancelar tercero
                if do_guardar:
                    GUARDAR_SELECTORS = [
                        "xpath=//*[@id='frmTerceroNatural:guardar']",
                        "#frmTerceroNatural\\:guardar",
                    ]
                    btn_found = False
                    for sel in GUARDAR_SELECTORS:
                        try:
                            await filtros_tab.wait_for_selector(sel, timeout=3000)
                            btn_found = True
                            break
                        except Exception:
                            continue

                    if btn_found:
                        _ = await safe_click(filtros_tab, GUARDAR_SELECTORS, "Botón frmTerceroNatural:guardar")
                        try:
                            await filtros_tab.evaluate("""
                                () => {
                                    const btn = document.getElementById('frmTerceroNatural:guardar');
                                    if (btn) { btn.scrollIntoView({block:'center'}); btn.click(); }
                                }
                            """)
                        except Exception:
                            pass
                        await wait_for_loader(filtros_tab, timeout=12000)
                    else:
                        print("[ERROR] Botón frmTerceroNatural:guardar no presente.", flush=True)

                if do_cancelar:
                    CANCELAR_SELECTORS = [
                        "xpath=//*[@id='frmTerceroNatural:cancelar']",
                        "#frmTerceroNatural\\:cancelar",
                    ]
                    cancel_found = False
                    for sel in CANCELAR_SELECTORS:
                        try:
                            await filtros_tab.wait_for_selector(sel, timeout=3000)
                            cancel_found = True
                            break
                        except Exception:
                            continue

                    if cancel_found:
                        _ = await safe_click(filtros_tab, CANCELAR_SELECTORS, "Botón frmTerceroNatural:cancelar")
                        try:
                            await filtros_tab.evaluate("""
                                () => {
                                    const btn = document.getElementById('frmTerceroNatural:cancelar');
                                    if (btn) { btn.scrollIntoView({block:'center'}); btn.click(); }
                                }
                            """)
                        except Exception:
                            pass
                        await wait_for_loader(filtros_tab, timeout=12000)
                    else:
                        print("[ERROR] Botón frmTerceroNatural:cancelar no presente.", flush=True)

                # tipoEnvio, claveLider, agePactado, siguiente
                TIPO_ENVIO_SEL = "#frmDatosFijos\\:componenteTercero\\:tipoEnvio"
                tipo_envio_val = CONFIG.get("tipo_envio", "PE")
                try:
                    _ = await select_richfaces_value(filtros_tab, TIPO_ENVIO_SEL, tipo_envio_val, wait_ajax_ms=1200)
                except Exception as e:
                    print(f"[ERROR] Error al seleccionar tipoEnvio -> {e}", flush=True)

                try:
                    await filtros_tab.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                except Exception:
                    pass

                CLAVE_LIDER_SEL = "#frmDatosFijos\\:claveLider"
                await safe_fill_and_tab(filtros_tab, CLAVE_LIDER_SEL, CONFIG["clave_lider"], "Clave Lider")
                await wait_for_loader(filtros_tab, timeout=8000)

                AGE_PACTADO_SEL = "#frmDatosFijos\\:agePactado"
                await safe_fill_and_tab(filtros_tab, AGE_PACTADO_SEL, CONFIG["age_pactado"], "Age Pactado")
                await wait_for_loader(filtros_tab, timeout=8000)

                CONTINUAR_SELECTORS = [
                    "#frmDatosFijos\\:siguiente",
                    "xpath=//*[@id='frmDatosFijos:siguiente']",
                    "input[value='Continuar »']"
                ]
                _ = await safe_click(filtros_tab, CONTINUAR_SELECTORS, "Botón frmDatosFijos:siguiente")
                await wait_for_loader(filtros_tab, timeout=12000)

                # formaDatosVariables:continuar + Localidad
                FDV_CONT_SELECTORS = [
                    "#formaDatosVariables\\:continuar",
                    "xpath=//*[@id='formaDatosVariables:continuar']",
                    "input[name='formaDatosVariables:continuar']",
                ]
                _ = await safe_click(filtros_tab, FDV_CONT_SELECTORS,
                                     "Botón formaDatosVariables:continuar (primer click)")
                await wait_for_loader(filtros_tab, timeout=12000)

                LOCALIDAD_LABEL_SEL = "xpath=//span[contains(@class,'dataLabel')][contains(normalize-space(),'Localidad')]"
                LOCALIDAD_INPUT_SEL = "xpath=//span[contains(@class,'dataLabel')][contains(normalize-space(),'Localidad')]/ancestor::tr[1]//input[@type='text']"

                try:
                    await filtros_tab.wait_for_selector(LOCALIDAD_LABEL_SEL, timeout=8000)
                    await filtros_tab.wait_for_selector(LOCALIDAD_INPUT_SEL, timeout=8000)
                    loc_input = filtros_tab.locator(LOCALIDAD_INPUT_SEL)
                    await loc_input.scroll_into_view_if_needed()
                    await loc_input.click()
                    await loc_input.fill("")
                    localidad_val = CONFIG.get("localidad", "BOGOTA")
                    await loc_input.type(localidad_val, delay=150)
                    await filtros_tab.press(LOCALIDAD_INPUT_SEL, "Enter")
                    await wait_for_loader(filtros_tab, timeout=10000)

                    # validaRiesgo
                    VALIDA_RIESGO_SELECTORS = [
                        "#frmRiesgosPoliza\\:validaRiesgo",
                        "xpath=//*[@id='frmRiesgosPoliza:validaRiesgo']",
                        "input[name='frmRiesgosPoliza:validaRiesgo']",
                        "input[value='Continuar »']",
                    ]
                    _ = await safe_click(
                        filtros_tab,
                        VALIDA_RIESGO_SELECTORS,
                        "Botón frmRiesgosPoliza:validaRiesgo (click normal)"
                    )
                    try:
                        await filtros_tab.evaluate("""
                            () => {
                                const btn = document.getElementById('frmRiesgosPoliza:validaRiesgo');
                                if (btn) { btn.scrollIntoView({block:'center'}); btn.click(); }
                            }
                        """)
                        await filtros_tab.evaluate("""
                            () => {
                                const btn = document.getElementById('frmRiesgosPoliza:validaRiesgo');
                                if (btn) { btn.scrollIntoView({block:'center'}); btn.click(); }
                            }
                        """)
                    except Exception:
                        pass

                    await wait_for_loader(filtros_tab, timeout=20000)

                    # ===============================
                    # Placa + datos de vehículo
                    # ===============================
                    try:
                        AUTO_PLACA_SEL = "#frmRiesgosPoliza\\:autoPlaca"
                        await filtros_tab.wait_for_selector(AUTO_PLACA_SEL, timeout=10000)
                        await filtros_tab.fill(AUTO_PLACA_SEL, CONFIG["placa"])
                        await filtros_tab.press(AUTO_PLACA_SEL, "Tab")

                        try:
                            await filtros_tab.wait_for_selector("#statusPanel_container",
                                                                state="visible",
                                                                timeout=10000)
                        except PlaywrightTimeoutError:
                            pass

                        await wait_for_loader(filtros_tab, timeout=20000)

                        # autoMarcaInput
                        MARCA_SEL = "#frmRiesgosPoliza\\:autoMarcaInput"
                        await filtros_tab.wait_for_selector(MARCA_SEL, timeout=8000)
                        marca_val = (await filtros_tab.input_value(MARCA_SEL)).strip()
                        print(f"[INFO] Valor de autoMarcaInput después de la placa: '{marca_val}'", flush=True)

                        if not marca_val:
                            print("[INFO] autoMarcaInput vacío → llenando datos manuales de vehículo…", flush=True)

                            # Marca manual
                            marca_manual = CONFIG.get("auto_marca_manual") or ""
                            if marca_manual:
                                marca_input = filtros_tab.locator(MARCA_SEL)
                                await marca_input.click()
                                await marca_input.fill("")
                                await marca_input.type(marca_manual, delay=80)
                                await filtros_tab.press(MARCA_SEL, "Enter")
                                await wait_for_loader(filtros_tab, timeout=15000)

                            # ==========================
                            # USO por TEXTO DEL OPTION
                            # ==========================
                            uso_text = CONFIG.get("auto_uso_value") or ""
                            if uso_text:
                                USO_SEL = "#frmRiesgosPoliza\\:autoUso"
                                print(f"[INFO] Intentando seleccionar autoUso por texto='{uso_text}'", flush=True)
                                ok_uso = await select_option_by_text(
                                    filtros_tab,
                                    USO_SEL,
                                    uso_text,
                                    "Uso del vehículo (autoUso)",
                                    wait_ajax_ms=1500
                                )
                                if not ok_uso:
                                    try:
                                        curr_uso = await filtros_tab.eval_on_selector(
                                            USO_SEL,
                                            "el => el.value"
                                        )
                                        print(
                                            f"[WARN] autoUso no se pudo fijar en '{uso_text}', value actual='{curr_uso}'",
                                            flush=True
                                        )
                                    except Exception as e:
                                        print(f"[ERROR] No se pudo leer valor actual de autoUso -> {e}", flush=True)
                            else:
                                print("[INFO] CONFIG['auto_uso_value'] vacío, no se selecciona uso.", flush=True)

                            # ==========================
                            # MODELO por TEXTO DEL OPTION
                            # ==========================
                            modelo_text = CONFIG.get("auto_modelo_value") or ""
                            if modelo_text:
                                MODELO_SEL = "#frmRiesgosPoliza\\:autoModelo"
                                print(f"[INFO] Intentando seleccionar autoModelo por texto='{modelo_text}'", flush=True)
                                ok_modelo = await select_option_by_text(
                                    filtros_tab,
                                    MODELO_SEL,
                                    modelo_text,
                                    "Modelo del vehículo (autoModelo)",
                                    wait_ajax_ms=1500
                                )
                                if not ok_modelo:
                                    try:
                                        curr_modelo = await filtros_tab.eval_on_selector(
                                            MODELO_SEL,
                                            "el => el.value"
                                        )
                                        print(
                                            f"[WARN] autoModelo no se pudo fijar en '{modelo_text}', value actual='{curr_modelo}'",
                                            flush=True
                                        )
                                    except Exception as e:
                                        print(f"[ERROR] No se pudo leer valor actual de autoModelo -> {e}", flush=True)
                            else:
                                print("[INFO] CONFIG['auto_modelo_value'] vacío, no se selecciona modelo.", flush=True)

                            # COLOR
                            color_val = CONFIG.get("auto_color") or ""
                            if color_val:
                                await safe_fill(
                                    filtros_tab,
                                    ["#frmRiesgosPoliza\\:autoColor",
                                     "xpath=//*[@id='frmRiesgosPoliza:autoColor']"],
                                    color_val,
                                    "Color"
                                )

                            # MOTOR
                            motor_val = CONFIG.get("auto_motor") or ""
                            if motor_val:
                                await safe_fill(
                                    filtros_tab,
                                    ["#frmRiesgosPoliza\\:autoMotor",
                                     "xpath=//*[@id='frmRiesgosPoliza:autoMotor']"],
                                    motor_val,
                                    "Número de motor"
                                )

                            # CHASIS
                            chasis_val = CONFIG.get("auto_chasis") or ""
                            if chasis_val:
                                await safe_fill(
                                    filtros_tab,
                                    ["#frmRiesgosPoliza\\:autoChasis",
                                     "xpath=//*[@id='frmRiesgosPoliza:autoChasis']"],
                                    chasis_val,
                                    "Número de chasis"
                                )

                            # KILOMETRAJE (si no está disabled)
                            km_val = CONFIG.get("auto_kilometraje") or ""
                            if km_val:
                                try:
                                    AUTO_KM_SEL = "#frmRiesgosPoliza\\:autoCeroKm"
                                    await filtros_tab.wait_for_selector(AUTO_KM_SEL, timeout=8000)
                                    is_disabled = await filtros_tab.eval_on_selector(
                                        AUTO_KM_SEL,
                                        "el => el.disabled === true"
                                    )
                                    if not is_disabled:
                                        await filtros_tab.fill(AUTO_KM_SEL, km_val)
                                    else:
                                        print("[INFO] autoCeroKm está disabled, no se escribe kilometraje.", flush=True)
                                except Exception as e:
                                    print(f"[WARN] No se pudo procesar autoCeroKm -> {e}", flush=True)

                            # SUMA ASEGURADA (después de modelo)
                            suma_val = CONFIG.get("auto_suma_asegurada") or ""
                            if suma_val:
                                SUMA_SEL = "#frmRiesgosPoliza\\:autoSumaAseg"
                                try:
                                    await filtros_tab.wait_for_selector(SUMA_SEL, timeout=8000)
                                    await filtros_tab.fill(SUMA_SEL, suma_val)
                                    await filtros_tab.press(SUMA_SEL, "Tab")
                                    await wait_for_loader(filtros_tab, timeout=20000)
                                except Exception as e:
                                    print(f"[WARN] No se pudo procesar autoSumaAseg -> {e}", flush=True)

                            # FECHA MATRÍCULA (si no está disabled)
                            fecha_mat_val = CONFIG.get("auto_fecha_matricula") or ""
                            if fecha_mat_val:
                                try:
                                    FECHA_MAT_SEL = "#frmRiesgosPoliza\\:autosFechaMatricula"
                                    await filtros_tab.wait_for_selector(FECHA_MAT_SEL, timeout=8000)
                                    is_disabled = await filtros_tab.eval_on_selector(
                                        FECHA_MAT_SEL,
                                        "el => el.disabled === true"
                                    )
                                    if not is_disabled:
                                        await filtros_tab.fill(FECHA_MAT_SEL, fecha_mat_val)
                                    else:
                                        print("[INFO] autosFechaMatricula está disabled, no se escribe fecha.", flush=True)
                                except Exception as e:
                                    print(f"[WARN] No se pudo procesar autosFechaMatricula -> {e}", flush=True)

                            # COBERTURA 43
                            cobertura_val = CONFIG.get("auto_cobertura_value") or ""
                            COB_SEL = "#frmRiesgosPoliza\\:autoCobertura"

                            if cobertura_val:
                                ok_cob = await select_richfaces_value(
                                    filtros_tab,
                                    COB_SEL,
                                    cobertura_val,
                                    wait_ajax_ms=1200,
                                )
                                if not ok_cob:
                                    print("[WARN] No se pudo seleccionar cobertura por value, intento por texto '43 -'",
                                          flush=True)
                                    try:
                                        await filtros_tab.evaluate("""
                                            () => {
                                                const s = document.querySelector('#frmRiesgosPoliza:autoCobertura');
                                                if (!s) return;
                                                const opts = [...s.options];
                                                const opt = opts.find(o => o.text.trim().startsWith('43 -'));
                                                if (opt) {
                                                    s.value = opt.value;
                                                    s.dispatchEvent(new Event('change', { bubbles: true }));
                                                    s.dispatchEvent(new Event('input',  { bubbles: true }));
                                                    s.blur && s.blur();
                                                }
                                            }
                                        """)
                                    except Exception as e:
                                        print(f"[ERROR] No se pudo seleccionar cobertura 43 -> {e}", flush=True)
                            else:
                                print("[WARN] auto_cobertura_value vacío en CONFIG.", flush=True)

                        else:
                            print("[INFO] autoMarcaInput ya tiene datos, no se llena manualmente.", flush=True)

                    except Exception as e:
                        print(f"[ERROR] No se pudo procesar autoPlaca / datos vehiculo -> {e}", flush=True)

                except Exception:
                    print("[INFO] No se encontró sección Localidad (quizá no aplica).", flush=True)

                # FIN
                success = True
                url_final = filtros_tab.url

                if CONFIG.get("keep_open", True):
                    try:
                        print(
                            "\n---\nScript finalizado. El navegador y las pestañas quedan ABIERTOS.\n"
                            "Presiona ENTER para terminar solo el PROCESO (el navegador NO se cierra automáticamente)…",
                            flush=True
                        )
                        await asyncio.get_event_loop().run_in_executor(None, input)
                    except Exception:
                        print("Entorno sin stdin. Manteniendo proceso vivo… (Ctrl+C para terminar)", flush=True)
                        await asyncio.sleep(24 * 3600)

    except Exception as e:
        err = str(e)
        pr(f"Error general: {err}", "ERROR")

    pr("Resumen JSON:")
    print(json.dumps({"success": success, "urlFinal": url_final, "error": err}, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    asyncio.run(main())
