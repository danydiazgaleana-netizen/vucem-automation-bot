"""
Script Principal de Orquestación - Automatización VUCEM
Autor: Daniela Diaz Galeana
Descripción: Coordina la carga de datos, checkpoints y ejecución
de la automatización en VUCEM. Menú interactivo que se repite después de cada modelo.
"""
import logging
import sys
import time
from config import Config
from data_manager import DataManager
from browser_automation import VUCEMAutomation, SessionExpiredException
from state_manager import StateManager
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

# Configuración de logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("vucem_execution.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

console = Console()
MAX_RETRIES_PER_MODEL = 3

def mostrar_menu(modelos, state_mgr):
    """Muestra el menú interactivo con modelos pendientes y opciones."""
    console.clear()
    console.rule("📋 VUCEM Automation Bot", style="bold blue")
    summary = state_mgr.get_summary()
    console.print(Panel(
        f"[bold green]✅ Éxitos (total):[/] {summary['processed']}   "
        f"[bold red]❌ Fallos (total):[/] {summary['failed']}   "
        f"[bold yellow]⏳ Pendientes:[/] {len(modelos)}",
        title="Estado del lote",
        border_style="cyan"
    ))

    if not modelos:
        console.print("[bold yellow]No hay modelos pendientes. ¡Todo procesado![/]")
        return False

    table = Table(title="Modelos pendientes", show_header=True, header_style="bold magenta")
    table.add_column("#", style="dim", width=4)
    table.add_column("Código", style="cyan", width=20)
    table.add_column("Nombre", style="white")
    for idx, m in enumerate(modelos, 1):
        table.add_row(str(idx), m['codigo'], m.get('nombre', m['codigo']))
    console.print(table)

    console.print("\n[bold]Opciones:[/]")
    console.print(f"  [cyan]1-{len(modelos)}[/]  Seleccionar modelo para procesar")
    console.print("  [yellow]R[/]  Reintentar modelos fallidos")
    console.print("  [green]E[/]  Exportar reporte Excel (modelos procesados con folio)")
    console.print("  [red]S[/]  Salir (guardar estado)")

    while True:
        entrada = input("\n👉 Selecciona una opción (número, 'R', 'E' o 'S'): ").strip().upper()
        if entrada in ('R', 'E', 'S'):
            return entrada
        if entrada.isdigit():
            idx = int(entrada)
            if 1 <= idx <= len(modelos):
                return idx
            else:
                console.print(f"[red]Número fuera de rango (1-{len(modelos)})[/]")
        else:
            # Buscar por nombre (coincidencia parcial)
            coincidencias = [m for m in modelos if entrada.lower() in m['nombre'].lower()]
            if len(coincidencias) == 1:
                idx = modelos.index(coincidencias[0]) + 1
                console.print(f"[green]Seleccionado: {coincidencias[0]['nombre']}[/]")
                return idx
            elif len(coincidencias) > 1:
                console.print("[yellow]Múltiples coincidencias, elige por número:[/]")
                for i, m in enumerate(coincidencias, 1):
                    console.print(f"  {i}. {m['nombre']}")
                try:
                    sub_idx = int(input("Número: "))
                    if 1 <= sub_idx <= len(coincidencias):
                        modelo_seleccionado = coincidencias[sub_idx - 1]
                        idx = modelos.index(modelo_seleccionado) + 1
                        return idx
                    else:
                        console.print("[red]Número fuera de rango.[/]")
                except ValueError:
                    console.print("[red]Entrada inválida.[/]")
            else:
                console.print("[red]No se encontró ningún modelo con ese nombre.[/]")

def procesar_modelo(bot, state_mgr, modelo, modo_automatico=False):
    """Procesa un modelo con reintentos y retorna (success, folio)."""
    codigo = modelo['codigo']
    for attempt in range(1, MAX_RETRIES_PER_MODEL + 1):
        try:
            folio = bot.process_model(modelo['data'], modo_automatico=modo_automatico)
            state_mgr.mark_completed(codigo, folio)
            logging.info(f"✅ Modelo {codigo} completado (intento {attempt}). Folio: {folio}")
            return True, folio
        except SessionExpiredException:
            logging.critical(f"🔴 Sesión expirada en {codigo}. No se puede continuar.")
            state_mgr.mark_failed(codigo, "Sesión expirada", traceback="")
            raise  # Propagar para detener el bucle
        except Exception as e:
            logging.warning(f"⚠️ Intento {attempt}/{MAX_RETRIES_PER_MODEL} falló: {e}")
            if attempt < MAX_RETRIES_PER_MODEL:
                wait = 2 ** attempt
                logging.info(f"⏳ Esperando {wait}s antes de reintentar...")
                time.sleep(wait)
            else:
                logging.error(f"❌ Modelo {codigo} falló después de {MAX_RETRIES_PER_MODEL} intentos.")
                state_mgr.mark_failed(codigo, str(e), traceback="")
                return False, None
    return False, None

def main():
    logging.info("🚀 Iniciando sistema automatizado de comercio exterior - VUCEM")

    # 1. Preparar datos
    data_mgr = DataManager()
    try:
        all_models = data_mgr.prepare_models()
    except Exception as e:
        logging.critical(f"❌ Preparación de datos fallida: {e}")
        sys.exit(1)

    if not all_models:
        logging.warning("⚠️ No hay modelos para procesar.")
        return

    # 2. Cargar estado
    state_mgr = StateManager()
    pending = state_mgr.get_pending(all_models)

    if not pending:
        logging.info("✅ Todos los modelos ya están procesados o en fallos.")
        summary = state_mgr.get_summary()
        logging.info(f"📊 Resumen: Éxitos: {summary['processed']}, Fallos: {summary['failed']}")
        return

    logging.info(f"📦 {len(pending)} modelos pendientes de {len(all_models)} totales.")

    # 3. Inicializar bot (solo una vez)
    bot = VUCEMAutomation()
    exitos_sesion = 0
    fallos_sesion = 0

    try:
        bot.login_with_manual_verification()
        bot.navigate_to_form()

        # Bucle interactivo: muestra menú después de cada modelo
        while True:
            # Refrescar pendientes
            pending = state_mgr.get_pending(all_models)
            if not pending:
                logging.info("✅ Todos los modelos pendientes han sido procesados.")
                break

            opcion = mostrar_menu(pending, state_mgr)

            if opcion == 'S':
                logging.info("🛑 Salida solicitada por el usuario.")
                break

            elif opcion == 'E':
                # Exportar reporte Excel
                excel_path = state_mgr.export_report_excel()
                if excel_path:
                    console.print(f"[green]✅ Reporte Excel generado: {excel_path}[/]")
                else:
                    console.print("[yellow]⚠️ No hay modelos completados para exportar.[/]")
                input("Presiona Enter para continuar...")
                continue

            elif opcion == 'R':
                fallidos = state_mgr.failures
                if not fallidos:
                    console.print("[yellow]No hay modelos fallidos para reintentar.[/]")
                    input("Presiona Enter para continuar...")
                    continue
                console.print(f"[cyan]Reintentando {len(fallidos)} modelos fallidos...[/]")
                for f in fallidos:
                    codigo = f['codigo']
                    modelo = next((m for m in all_models if m['codigo'] == codigo), None)
                    if modelo is None:
                        logging.error(f"⚠️ Modelo con código {codigo} no encontrado en la lista original.")
                        continue
                    state_mgr.failures = [x for x in state_mgr.failures if x['codigo'] != codigo]
                    state_mgr._save_failures()
                    success, folio = procesar_modelo(bot, state_mgr, modelo, modo_automatico=False)
                    if success:
                        exitos_sesion += 1
                    else:
                        fallos_sesion += 1
                input("Presiona Enter para continuar...")
                continue

            else:
                # Selección de modelo por número
                idx = int(opcion) - 1
                modelo = pending[idx]
                success, folio = procesar_modelo(bot, state_mgr, modelo, modo_automatico=False)
                if success:
                    exitos_sesion += 1
                else:
                    fallos_sesion += 1
                input("Presiona Enter para volver al menú...")

    except KeyboardInterrupt:
        logging.info("🛑 Interrupción por usuario. Estado guardado.")
    except SessionExpiredException:
        logging.critical("🔴 Sesión expirada. El proceso se detiene.")
    except Exception as e:
        logging.critical(f"❌ Error crítico: {e}")
    finally:
        bot.close()
        summary = state_mgr.get_summary()
        if hasattr(state_mgr, 'generar_reporte_html'):
            state_mgr.generar_reporte_html()
        logging.info("=" * 60)
        logging.info("🏁 Procesamiento finalizado.")
        logging.info(f"📊 Resumen de esta sesión:")
        logging.info(f"   ✅ Procesados en esta sesión: {exitos_sesion}")
        logging.info(f"   ❌ Fallos en esta sesión: {fallos_sesion}")
        logging.info(f"📊 Acumulado total:")
        logging.info(f"   ✅ Total procesados: {summary['processed']}")
        logging.info(f"   ❌ Total fallidos: {summary['failed']}")
        if summary['failed'] > 0:
            logging.info(f"   📄 Revisa 'failed_models.json' para detalles.")
        logging.info("=" * 60)

if __name__ == "__main__":
    main()