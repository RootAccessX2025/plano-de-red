import json
import random

def generar_puestos_izquierda():
    """Genera un archivo JSON con los puestos definidos en el index.html."""
    
    # Definir servicios, VLANs y colores basados en la imagen
    servicios = {
        "CREDIFINANCIERA": {"vlan": 10, "color": "#6aa84f"},  # Verde
        "ENDESA": {"vlan": 20, "color": "#e69138"},  # Naranja
        "CITAS": {"vlan": 30, "color": "#f1c232"},  # Amarillo
        "CUADRO MEDICO": {"vlan": 40, "color": "#a64d79"},  # Morado
        "EURONA": {"vlan": 50, "color": "#674ea7"}  # Púrpura
    }


    # Puestos extraídos del index.html
    puestos_lista = [
        ["0182", "0183", "0184"],
        ["PASILLO"],
        ["0200", "0201", "0202", "0203", "0204", "0205", "0206", "0207", "0208", "0209", "0210", "0211", "0212", "019"],
        ["0226", "0227", "0228", "0229", "0230", "0231", "0232", "0233", "0234", "0235", "0236", "0237", "0238", "020"],
        ["PASILLO"],
        ["0252", "0253", "0254", "0255", "0256", "0257", "0258", "0259", "0260", "COLUMNA", "0261", "0262", "021"],
        ["0275", "0276", "0277", "0278", "0279", "0280", "0281", "0282", "0283", "COLUMNA", "0284", "0285", "022"],
        ["PASILLO"],
        ["0298", "0299", "0300", "0301", "0302", "0303", "0304", "0305", "0306", "0307", "0308", "0309", "0310", "023"],
        ["0324", "0325", "0326", "0327", "0328", "0329", "0330", "0331", "0332", "0333", "0334", "0335", "0336", "024"],
        ["PASILLO"],
        ["0350", "0351", "0352", "0353", "0354", "0355", "0356", "0357", "0358", "COLUMNA", "0359", "0360", ""],
        ["0373", "0374", "0375", "0376", "0377", "0378", "0379", "0380", "0381", "COLUMNA", "0382", "0383", ""],
        ["PASILLO"],
        ["0396", "0397", "0398", "0399", "0400", "0401", "0402", "0403", "0404", "0405", "0406", "0407", "0408", "027"],
        ["0422", "0423", "0424", "0425", "0426", "0427", "0428", "0429", "0430", "0431", "0432", "0433", "0434", "028"],
        ["PASILLO"],
        ["0448", "0449", "0450", "0451", "0452", "0453", "0454", "0455", "0456", "COLUMNA", "0457", "0458", "029"],
        ["0471", "0472", "0473", "0474", "0475", "0476", "0477", "0478", "0479", "COLUMNA", "0480", "0481", "030"],
        ["PASILLO"],
        ["0494", "0495", "0496", "0497", "0498", "0499", "0500", "0501", "0502", "0503", "0504", "0505", "0506", "0507"],
    ]

    puestos = []
    for fila in puestos_lista:
        fila_puestos = []
        for puesto in fila:
            if puesto == "COLUMNA":
                fila_puestos.append({"id": "COLUMNA", "label": "", "color": "#000000"})
            elif puesto == "PASILLO":
                fila_puestos.append({"id": "PASILLO", "label": "", "color": "#cccccc"})
            else:
                servicio = random.choice(list(servicios.keys()))
                info_servicio = servicios[servicio]
                fila_puestos.append({
                    "id": f"PUESTO_{puesto}",
                    "label": puesto,
                    "puerto_red": f"GigabitEthernet0/{random.randint(1, 48)}",
                    "vlan": info_servicio["vlan"],
                    "color": info_servicio["color"],
                    "switch": f"Switch-{random.randint(1, 10)}",
                    
                })
        puestos.append(fila_puestos)

    # Guardar en JSON
    with open("puestos_red.json", "w") as f:
        json.dump(puestos, f, indent=4)

    return puestos

if __name__ == "__main__":
    puestos = generar_puestos_izquierda()
    print("✅ Mapa de la parte izquierda generado correctamente.")