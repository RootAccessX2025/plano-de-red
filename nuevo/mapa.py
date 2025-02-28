import json

def generar_puestos_red():
    """Genera un archivo JSON con los puestos de trabajo organizados correctamente con canaleta y pasillo"""

    # Definir servicios y VLANs
    servicios = {
        "CREDIFINANCIERA": {"vlan": 10, "color": "green"},
        "ENDESA": {"vlan": 20, "color": "orange"},
        "CITAS": {"vlan": 30, "color": "yellow"},
        "CUADRO MEDICO": {"vlan": 40, "color": "purple"},
        "EURONA": {"vlan": 50, "color": "blue"}
    }

    # Configuración de la cuadrícula con pasillo y canaletas
    filas = 10  # Filas de puestos
    columnas = 18  # Columnas en total (9 a la izquierda, 9 a la derecha del pasillo)
    espaciado_x = 120  # Espaciado horizontal
    espaciado_y = 80  # Espaciado vertical
    pasillo_ancho = 300  # Espacio en el centro para el pasillo
    canaleta_y_offset = 40  # Altura de la canaleta respecto a los puestos

    puestos = []
    canaletas = []
    num_puesto = 1

    # Generar primera sección de puestos (lado izquierdo del pasillo)
    for fila in range(filas):
        for col in range(columnas // 2):  # Primera mitad antes del pasillo
            servicio = list(servicios.keys())[fila % len(servicios)]
            info_servicio = servicios[servicio]

            x = col * espaciado_x  # Mantener los puestos juntos
            y = fila * espaciado_y

            puestos.append({
                "id": f"PC_{num_puesto}",
                "label": f"PC {num_puesto}",
                "servicio": servicio,
                "puerto_red": f"GigabitEthernet0/{num_puesto}",
                "vlan": info_servicio["vlan"],
                "color": info_servicio["color"],
                "x": x,
                "y": y
            })

            # Crear nodo de canaleta para agrupar los puertos en la misma fila
            if col == 0:
                canaletas.append({
                    "id": f"CANALETA_{fila}",
                    "label": f"Canaleta {fila + 1}",
                    "shape": "box",
                    "color": "gray",
                    "x": x + (espaciado_x * (columnas // 4)),  # Centrar la canaleta
                    "y": y + canaleta_y_offset
                })

            num_puesto += 1

    # Generar segunda sección de puestos (lado derecho del pasillo)
    for fila in range(filas):
        for col in range(columnas // 2, columnas):  # Segunda mitad después del pasillo
            servicio = list(servicios.keys())[fila % len(servicios)]
            info_servicio = servicios[servicio]

            x = (col * espaciado_x) + pasillo_ancho  # Desplazar los puestos a la derecha del pasillo
            y = fila * espaciado_y

            puestos.append({
                "id": f"PC_{num_puesto}",
                "label": f"PC {num_puesto}",
                "servicio": servicio,
                "puerto_red": f"GigabitEthernet0/{num_puesto}",
                "vlan": info_servicio["vlan"],
                "color": info_servicio["color"],
                "x": x,
                "y": y
            })

            # Crear nodo de canaleta en la misma fila
            if col == columnas // 2:
                canaletas.append({
                    "id": f"CANALETA_{fila}_D",
                    "label": f"Canaleta {fila + 1}",
                    "shape": "box",
                    "color": "gray",
                    "x": x - (espaciado_x * (columnas // 4)),  # Centrar la canaleta en la derecha
                    "y": y + canaleta_y_offset
                })

            num_puesto += 1

    # Guardar en JSON
    with open("puestos_red.json", "w") as f:
        json.dump(puestos + canaletas, f, indent=4)

    print("✅ Puestos organizados con canaleta y pasillo correctamente.")

if __name__ == "__main__":
    generar_puestos_red()
