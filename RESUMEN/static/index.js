document.addEventListener("DOMContentLoaded", () => {
    console.log("Script unificado listo para procesar solicitudes");

    document.addEventListener('click', (e) => {
        const boton = e.target.closest('.btn-procesar');
        if (!boton) return;

        e.preventDefault();

        // Obtener la clave del módulo asignada en data-modulo
        const modulo = boton.getAttribute('data-modulo');
        
        // Elementos correspondientes a la sección clickeada
        const inputArchivo = document.getElementById(`file-${modulo}`);
        const contenedorResultado = document.getElementById(`resultado-${modulo}`);
        const archivo = inputArchivo ? inputArchivo.files[0] : null;

        if (!archivo) {
            alert("Por favor selecciona un archivo Excel primero.");
            return;
        }

        const datos = new FormData();
        datos.append("file", archivo);

        if (contenedorResultado) {
            contenedorResultado.style.display = "block";
            contenedorResultado.innerHTML = `<h3 style="color: #2d6cdf;">Procesando '${archivo.name}', por favor espera...</h3>`;
        }

        const rutaFlask = `/upload/${modulo}`;
        console.log(`Enviando archivo a la ruta: ${rutaFlask}`);

        fetch(rutaFlask, {
            method: "POST",
            body: datos
        })
        .then(respuesta => {
            if (!respuesta.ok) {
                throw new Error(`Error en el servidor Flask. Estado HTTP: ${respuesta.status}`);
            }
            return respuesta.text();
        })
        .then(html => {
            if (contenedorResultado) {
                contenedorResultado.innerHTML = html;
            }
        })
        .catch(error => {
            console.error("Error al procesar:", error);
            if (contenedorResultado) {
                contenedorResultado.innerHTML = `
                    <h3 style="color:red;">Error al procesar el archivo</h3>
                    <p style="color:gray;">Detalle: ${error.message}</p>
                `;
            }
        });
    });
});