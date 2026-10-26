export const metadata = {
  title: "Proyecto Educativo en los Uros — Puno",
  description: "Control de documentación del Proyecto Educativo en los Uros - Puno",
};

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body style={{ margin: 0, padding: 0, background: "#010708" }}>
        {children}
      </body>
    </html>
  );
}
