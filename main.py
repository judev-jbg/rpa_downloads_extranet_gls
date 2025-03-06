"""
Punto de entrada principal para la aplicación RPA de envíos GLS.
Implementado con Selenium WebDriver para mayor robustez y fiabilidad.
"""
import logging
from rpa import rpa_shipments

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("rpa_shipments.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("Toolstock-GLS RPA Main")

def run_rpa():
    """Ejecuta el proceso RPA de envíos."""
    try:
        logger.info("Iniciando proceso RPA para envíos GLS")
        
        if rpa_shipments():
            logger.info("Proceso RPA de envíos finalizado con éxito")
            return True
        else:
            logger.error("Error en el proceso RPA de envíos")
            return False
    except Exception as e:
        logger.error(f"Error crítico en la ejecución del RPA: {e}")
        return False

if __name__ == "__main__":
    run_rpa()