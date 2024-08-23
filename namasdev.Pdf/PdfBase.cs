using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace namasdev.Pdf
{
    public abstract class PdfBase
    {
        private readonly Guid _id;
        private string _pathDirectorioImagenes;
        private int _imagenTemporalId;

        public PdfBase(string titulo)
        {
            if (String.IsNullOrWhiteSpace(titulo))
                throw new ArgumentNullException("titulo");

            _id = Guid.NewGuid();

            Titulo = titulo;
            NombreArchivo = GenerarNombreArchivo();
        }

        #region Propiedades

        #region Publicos

        public string Titulo { get; set; }
        public string NombreArchivo { get; private set; }

        #endregion

        #region Protegidos

        protected Section Section { get; private set; }
        protected Document Document { get; private set; }

        #endregion Protegidos

        #endregion Propiedades

        #region Metodos

        #region Publicos

        public void Guardar(string path)
        {
            try
            {
                var bytes = pGenerarYDevolverBytes();
                File.WriteAllBytes(path, bytes);
            }
            finally
            {
                LimpiarImagenesTemporales();
            }
        }

        public void Guardar(Stream stream)
        {
            try
            {
                GenerarYGuardarEnStream(stream);
            }
            finally
            {
                LimpiarImagenesTemporales();
            }
        }

        public byte[] GenerarYDevolverBytes()
        {
            try
            {
                return pGenerarYDevolverBytes();
            }
            finally
            {
                LimpiarImagenesTemporales();
            }
        }

        public void AddSection()
        {
            Section = Document.AddSection();
        }

        #endregion Publicos

        #region Protegidos

        protected void Generar()
        {
            Document = new Document();
            Document.Info.Title = Titulo;
            //Document.Info.Author = autor;

            Section = Document.AddSection();

            DefinirEstilosPagina();

            GenerarEncabezado();
            GenerarPie();
            GenerarContenido();
        }

        protected abstract string ObtenerPathDirectorioTemporalImagenes();
        protected abstract void DefinirEstilosPagina();
        protected abstract void GenerarEncabezado();
        protected abstract void GenerarPie();
        protected abstract void GenerarContenido();

        /// <summary>
        /// Descarga la imagen desde la URI especificada y la guarda temporalmente.
        /// Finalmente, devuelve el nombre de la imagen temporal.
        /// </summary>
        /// <param name="imagenUri"></param>
        /// <param name="imagenExtension"></param>
        /// <returns>Devuelve el nombre de la imagen temporal.</returns>
        protected string GuardarImagenTemporal(Uri imagenUri,
            string imagenExtension = null)
        {
            CrearCarpetaImagenesTemporalesSiNoExiste();

            imagenExtension = imagenExtension ?? Path.GetExtension(imagenUri.LocalPath);

            var nombreImagen = GenerarNombreImagen(imagenExtension);
            var pathImagenTemporal = GenerarPathImagenTemporal(nombreImagen);

            using (var webClient = new WebClient())
            {
                try
                {
                    webClient.DownloadFile(imagenUri, pathImagenTemporal);
                }
                catch (Exception)
                {
                    //  no hago nada
                }
            }

            return nombreImagen;
        }

        protected void AgregarSeparador(
            Section seccion,
            Unit ancho, Unit alto,
            bool conLinea = false, Color? color = null)
        {
            if (seccion == null)
            {
                throw new ArgumentNullException("seccion");
            }

            GenerarTablaSeparador(Section.AddTable(), ancho, alto,
                conLinea: conLinea,
                color: color);
        }

        protected void AgregarSeparador(
            HeaderFooter encabezadoOPie, 
            Unit ancho, Unit alto,
            bool conLinea = false, Color? color = null)
        {
            if (encabezadoOPie == null)
            {
                throw new ArgumentNullException("encabezadoOPie");
            }

            GenerarTablaSeparador(encabezadoOPie.AddTable(), ancho, alto,
                conLinea: conLinea,
                color: color);
        }

        protected void GenerarTablaSeparador(Table tablaSeparador, Unit ancho, Unit alto,
            bool conLinea = false, Color? color = null)
        {
            if (tablaSeparador == null)
            {
                throw new ArgumentNullException("tablaSeparador");
            }

            tablaSeparador.AddColumn(ancho);
            tablaSeparador.Rows.HeightRule = RowHeightRule.Exactly;
            tablaSeparador.Rows.Height = alto;

            var f1 = tablaSeparador.AddRow();

            if (conLinea)
            {
                f1.Borders.Bottom.Color = color ?? Colors.Black;
            }

            tablaSeparador.AddRow();
        }

        #endregion Protegidos

        #region Privados

        private string GenerarNombreArchivo()
        {
            return String.Format("{0}.pdf", Titulo);
        }

        private void GenerarYGuardarEnStream(Stream stream)
        {
            Generar();

            var renderer = Renderizar();
            renderer.Save(stream, false);
        }

        private MigraDoc.Rendering.PdfDocumentRenderer Renderizar()
        {
            if (Document == null)
                throw new ArgumentNullException("Document");

            var renderer = new MigraDoc.Rendering.PdfDocumentRenderer();
            renderer.Document = Document;
            renderer.RenderDocument();

            return renderer;
        }

        private byte[] pGenerarYDevolverBytes()
        {
            using (var stream = new MemoryStream())
            {
                GenerarYGuardarEnStream(stream);

                return stream.ToArray();
            }
        }

        private void CrearCarpetaImagenesTemporalesSiNoExiste()
        {
            if (!String.IsNullOrWhiteSpace(_pathDirectorioImagenes))
            {
                return;
            }

            var pathDirectorioTemporalImagenes = ObtenerPathDirectorioTemporalImagenes();
            if (String.IsNullOrWhiteSpace(pathDirectorioTemporalImagenes))
            {
                throw new ApplicationException("No ha especificado el path del directorio donde se guardarán las imágenes temporales.");
            }

            var path = Path.Combine(pathDirectorioTemporalImagenes, _id.ToString());
            Directory.CreateDirectory(path);

            Document.ImagePath = _pathDirectorioImagenes = path;

            _imagenTemporalId = 1;
        }

        private void LimpiarImagenesTemporales()
        {
            if (!String.IsNullOrWhiteSpace(_pathDirectorioImagenes))
            {
                try
                {
                    Directory.Delete(_pathDirectorioImagenes, true);
                }
                catch (Exception ex)
                {
                    // TODO (ML): registrar error
                }
            }

            Document.ImagePath = _pathDirectorioImagenes = null;
        }

        private string GenerarNombreImagen(string imagenExtension)
        {
            if (String.IsNullOrWhiteSpace(imagenExtension))
                throw new ArgumentNullException("imagenExtension");

            imagenExtension = imagenExtension.TrimStart('.');

            var nombre = String.Format("{0:D10}", _imagenTemporalId);
            _imagenTemporalId++;

            return String.Format("{0}.{1}", nombre, imagenExtension);
        }

        private string GenerarPathImagenTemporal(string nombreImagen)
        {
            return Path.Combine(_pathDirectorioImagenes, nombreImagen);
        }

        #endregion Privados

        #endregion Metodos

        #region Clases internas

        public class ColumnaFormato
        {
            public Unit Width { get; set; }
            public ParagraphFormat Format { get; set; }
            public Borders Borders { get; set; }
            public Shading Shading { get; set; }
            public string Style { get; set; }

            public void Aplicar(Column columna)
            {
                if (columna == null)
                    throw new ArgumentNullException("columna");

                if (Format != null)
                    columna.Format = Format.Clone();

                if (Borders != null)
                    columna.Borders = Borders.Clone();

                if (Shading != null)
                    columna.Shading = Shading.Clone();

                if (!String.IsNullOrWhiteSpace(Style))
                    columna.Style = Style;
            }
        }

        public class FilaFormato
        {
            public Borders Borders { get; set; }
            public ParagraphFormat Format { get; set; }
            public Unit? Height { get; set; }
            public RowHeightRule? HeightRule { get; set; }
            public Shading Shading { get; set; }
            public VerticalAlignment? VerticalAlignment { get; set; }
            public string Style { get; set; }

            public void Aplicar(Row fila)
            {
                if (fila == null)
                    throw new ArgumentNullException("fila");

                if (Borders != null)
                    fila.Borders = Borders.Clone();

                if (Format != null)
                    fila.Format = Format.Clone();

                if (Height.HasValue)
                    fila.Height = Height.Value;

                if (HeightRule.HasValue)
                    fila.HeightRule = HeightRule.Value;

                if (Shading != null)
                    fila.Shading = Shading.Clone();

                if (VerticalAlignment.HasValue)
                    fila.VerticalAlignment = VerticalAlignment.Value;

                if (!String.IsNullOrWhiteSpace(Style))
                    fila.Style = Style;
            }
        }

        public class TablaFormato
        {
            public Borders Borders { get; set; }
            public ParagraphFormat Format { get; set; }
            public Shading Shading { get; set; }
            public string Style { get; set; }

            public void Aplicar(Table table)
            {
                if (table == null)
                    throw new ArgumentNullException("table");

                if (Borders != null)
                    table.Borders = Borders.Clone();

                if (Format != null)
                    table.Format = Format.Clone();

                if (Shading != null)
                    table.Shading = Shading.Clone();

                if (!String.IsNullOrWhiteSpace(Style))
                    table.Style = Style;
            }
        }

        public abstract class TablaBase
        {
            protected TablaBase(Table table)
            {
                if (table == null)
                    throw new ArgumentNullException("table");

                Table = table;
            }

            protected Table Table { get; set; }

            private IEnumerable<ColumnaFormato> FormatoColumnas { get; set; }
            private FilaFormato FormatoFilas { get; set; }

            protected void Inicializar(IEnumerable<ColumnaFormato> columnasFormato, TablaFormato tablaFormato = null, FilaFormato filasFormato = null)
            {
                if (columnasFormato == null || columnasFormato.Count() == 0)
                    throw new ArgumentNullException("columnasFormato");

                FormatoColumnas = columnasFormato;
                FormatoFilas = filasFormato;

                if (tablaFormato != null)
                    tablaFormato.Aplicar(Table);

                foreach (var columnaFormato in columnasFormato)
                {
                    var columna = Table.AddColumn(columnaFormato.Width);
                    columnaFormato.Aplicar(columna);
                }
            }

            protected Row AgregarFila()
            {
                var fila = Table.AddRow();

                if (FormatoFilas != null)
                    FormatoFilas.Aplicar(fila);

                return fila;
            }

            protected void AplicarBordeSoloATabla(Edge edge, BorderStyle borderStyle, Unit unit, Color color)
            {
                Table.SetEdge(0, 0, Table.Columns.Count, Table.Rows.Count, edge, borderStyle, unit, color);
            }

            protected void MantenerFilasEnMismaPagina()
            {
                var cantFilas = Table.Rows.Count;
                if (cantFilas > 0)
                {
                    Table.Rows[0].KeepWith = cantFilas - 1;
                }
            }
        }

        #endregion Clases internas
    }
}
