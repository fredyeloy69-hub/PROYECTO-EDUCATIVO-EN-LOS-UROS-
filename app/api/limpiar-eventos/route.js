import { NextResponse } from "next/server";
import { adminDb } from "@/lib/firebaseAdmin";

export const maxDuration = 60;

// RESTAURA los eventos que cambiamos a "No identificado" (31/10/2026 18:30-18:45 Lima)
// y les devuelve el nombre que tenían: "<nombre> (última persona que lo subió/la modificó)".
// Uso:
//   /api/limpiar-eventos            -> solo MUESTRA cuántos restauraría
//   /api/limpiar-eventos?aplicar=1  -> aplica
// BORRAR este archivo después de usarlo.
export async function GET(request) {
  const url = new URL(request.url);
  const aplicar = url.searchParams.get("aplicar") === "1";
  try {
    const desde = new Date("2026-10-31T23:30:00Z");
    const hasta = new Date("2026-10-31T23:45:00Z");
    const snap = await adminDb
      .collection("eventos")
      .where("timestamp", ">=", desde)
      .where("timestamp", "<", hasta)
      .get();

    const cambios = [];
    for (const d of snap.docs) {
      const e = d.data();
      if (e.tipo !== "archivo_borrado" && e.tipo !== "carpeta_borrada") continue;
      if (e.usuario !== "No identificado" || !e.ultimoModificador) continue;
      const sufijo = e.tipo === "archivo_borrado" ? "lo subió" : "la modificó";
      cambios.push({
        ref: d.ref,
        item: e.item,
        despues: `${e.ultimoModificador} (última persona que ${sufijo})`,
      });
    }

    if (aplicar) {
      for (let i = 0; i < cambios.length; i += 100) {
        const b = adminDb.batch();
        for (const c of cambios.slice(i, i + 100)) b.update(c.ref, { usuario: c.despues });
        await b.commit();
      }
    }
    return NextResponse.json({
      ok: true,
      modo: aplicar ? "RESTAURADO" : "SOLO VISTA PREVIA (agrega ?aplicar=1 para restaurar)",
      eventosEnLaVentana: snap.size,
      eventosAReestaurar: cambios.length,
      ejemplos: cambios.slice(0, 3).map((c) => ({ item: c.item, despues: c.despues })),
    });
  } catch (err) {
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
