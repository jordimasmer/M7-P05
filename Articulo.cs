using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M7_P05
{
    class Articulo
    {
        private String nombre;
        private float precioUnidad;
        
        public Articulo(String nombre, float precio) {
            this.nombre = nombre;
            this.precioUnidad = precio;
        }

        public String gsNombre { 
            get { return nombre; } 
            set { nombre = value; } 
        }
        public float gsPrecio { 
            get { return precioUnidad; } 
            set { precioUnidad = value;} 
        }
    }
}
