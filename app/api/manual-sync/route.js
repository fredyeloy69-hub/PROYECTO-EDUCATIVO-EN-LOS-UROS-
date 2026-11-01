import { NextResponse } from "next/server";
import { adminDb } from "@/lib/firebaseAdmin";

export const maxDuration = 60;

// LIMPIEZA PUNTUAL: los eventos de borrado del sync de las 18:36 del 31/10/2026 se
// guardaron con el nombre de la última persona que subió el archivo, como si ella lo
// hubiera borrado. Esta ruta los deja como "No identificado" y conserva ese nombre en
// "ultimoModificador" (igual que hace ahora el sistema).
//
// Uso:
//   /api/limpiar-eventos            -> solo MUESTRA cuántos cambiaría (no toca nada)
//   /api/limpiar-eventos?aplicar=1  -> aplica el cambio
// BORRAR este archivo después de usarlo.
export async function GET(request) {
  const url = new URL(request.url);
  const aplicar = url.searchParams.get("aplicar") === "1";
  try {
    // 18:30 a 18:45 hora de Lima del 31/10/2026 (= 23:30 a 23:45 UTC)
    const desde = new Date("2026-10-31T23:30:00Z");
    const hasta = new Date("2026-10-31T23:45:00Z");
    const snap = await adminDb
      .collection("eventos")
      .where("timestamp", ">=", desde)
      .where("timestamp", "<", hasta)
      .get();

    const re = /^(.*) \(última persona que (?:lo subió|la modificó)\)$/;
    const cambios = [];
    for (const d of snap.docs) {
      const e = d.data();
      if (e.tipo !== "archivo_borrado" && e.tipo !== "carpeta_borrada") continue;
      const m = re.exec(String(e.usuario || ""));
      if (!m) continue;
      cambios.push({ ref: d.ref, item: e.item, antes: e.usuario, nombre: m[1] });
    }

    if (aplicar) {
      for (let i = 0; i < cambios.length; i += 100) {
        const b = adminDb.batch();
        for (const c of cambios.slice(i, i + 100)) {
          b.update(c.ref, { usuario: "No identificado", ultimoModificador: c.nombre });
        }
        await b.commit();
      }
    }
    return NextResponse.json({
      ok: true,
      modo: aplicar ? "APLICADO" : "SOLO VISTA PREVIA (agrega ?aplicar=1 para aplicar)",
      eventosEnLaVentana: snap.size,
      eventosAfectados: cambios.length,
      ejemplos: cambios.slice(0, 5).map((c) => ({ item: c.item, antes: c.antes })),
    });
  } catch (err) {
    console.error("Error limpiando eventos:", err);
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
