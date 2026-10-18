package com.office.offline;

import android.content.ContentResolver;
import android.content.ContentValues;
import android.content.Context;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.print.PrintAttributes;
import android.print.PrintDocumentAdapter;
import android.print.PrintManager;
import android.provider.MediaStore;
import android.util.Base64;
import android.webkit.MimeTypeMap;
import android.webkit.WebView;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Locale;

/**
 * Plugin native kecil:
 *  - save():  menyimpan berkas (base64) ke folder Download perangkat (MediaStore di Android 10+).
 *  - print(): membuka dialog cetak Android untuk halaman WebView (bisa "Simpan sebagai PDF").
 */
@CapacitorPlugin(name = "FileSaver")
public class FileSaverPlugin extends Plugin {

    @PluginMethod
    public void save(PluginCall call) {
        String name = call.getString("name", "berkas");
        String data = call.getString("data", "");
        String mime = call.getString("mime", "application/octet-stream");
        if (data == null || data.isEmpty()) {
            call.reject("Data kosong");
            return;
        }
        name = name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        if (name.isEmpty()) name = "berkas";
        mime = guessMime(name, mime);

        final String fName = name;
        final String fMime = mime;
        final byte[] bytes;
        try {
            bytes = Base64.decode(data, Base64.DEFAULT);
        } catch (IllegalArgumentException e) {
            call.reject("Data tidak valid");
            return;
        }

        // Tulis di thread terpisah supaya UI tidak macet untuk berkas besar.
        new Thread(() -> {
            try {
                Context ctx = getContext();
                JSObject res = new JSObject();
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                    ContentResolver cr = ctx.getContentResolver();
                    ContentValues v = new ContentValues();
                    v.put(MediaStore.MediaColumns.DISPLAY_NAME, fName);
                    v.put(MediaStore.MediaColumns.MIME_TYPE, fMime);
                    v.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS);
                    Uri uri = cr.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, v);
                    if (uri == null) throw new Exception("Tidak bisa membuat berkas di Download");
                    try (OutputStream os = cr.openOutputStream(uri)) {
                        if (os == null) throw new Exception("Tidak bisa menulis berkas");
                        os.write(bytes);
                    }
                    res.put("name", fName);
                    res.put("uri", uri.toString());
                } else {
                    // Android 9 ke bawah: simpan di folder Download milik aplikasi (tanpa izin tambahan).
                    File dir = ctx.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                    if (dir == null) dir = ctx.getFilesDir();
                    if (!dir.exists()) dir.mkdirs();
                    File f = uniqueFile(dir, fName);
                    try (FileOutputStream os = new FileOutputStream(f)) {
                        os.write(bytes);
                    }
                    res.put("name", f.getName());
                    res.put("uri", f.getAbsolutePath());
                }
                call.resolve(res);
            } catch (Exception e) {
                call.reject(e.getMessage() != null ? e.getMessage() : "Gagal menyimpan", e);
            }
        }).start();
    }

    @PluginMethod
    public void print(PluginCall call) {
        final String title = call.getString("title", "Dokumen");
        getActivity().runOnUiThread(() -> {
            try {
                WebView wv = getBridge().getWebView();
                PrintManager pm = (PrintManager) getActivity().getSystemService(Context.PRINT_SERVICE);
                PrintDocumentAdapter adapter = wv.createPrintDocumentAdapter(title);
                pm.print(title, adapter, new PrintAttributes.Builder().build());
                call.resolve();
            } catch (Exception e) {
                call.reject("Gagal membuka dialog cetak", e);
            }
        });
    }

    private static String guessMime(String name, String given) {
        if (given != null && !given.isEmpty() && !given.equals("application/octet-stream")) return given;
        int i = name.lastIndexOf('.');
        if (i >= 0) {
            String ext = name.substring(i + 1).toLowerCase(Locale.ROOT);
            switch (ext) {
                case "docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case "xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                case "pdf":  return "application/pdf";
                case "csv":  return "text/csv";
                case "html": case "htm": return "text/html";
                case "txt":  return "text/plain";
            }
            String m = MimeTypeMap.getSingleton().getMimeTypeFromExtension(ext);
            if (m != null) return m;
        }
        return "application/octet-stream";
    }

    private static File uniqueFile(File dir, String name) {
        File f = new File(dir, name);
        if (!f.exists()) return f;
        int dot = name.lastIndexOf('.');
        String base = dot > 0 ? name.substring(0, dot) : name;
        String ext = dot > 0 ? name.substring(dot) : "";
        for (int n = 1; n < 1000; n++) {
            f = new File(dir, base + " (" + n + ")" + ext);
            if (!f.exists()) return f;
        }
        return f;
    }
}
