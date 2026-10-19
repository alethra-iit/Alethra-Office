package com.office.offline;

import android.Manifest;
import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.ClipData;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.content.pm.ShortcutInfo;
import android.content.pm.ShortcutManager;
import android.database.Cursor;
import android.graphics.drawable.Icon;
import android.media.MediaScannerConnection;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.provider.OpenableColumns;
import android.provider.Settings;
import android.util.Base64;

import androidx.core.content.FileProvider;

import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayDeque;
import java.util.Locale;

/**
 * Plugin "Files": daftar & kelola dokumen di penyimpanan perangkat.
 * Membutuhkan izin "Akses semua berkas" (Android 11+) atau izin penyimpanan (Android 10 ke bawah).
 */
@CapacitorPlugin(name = "Files")
public class FilesPlugin extends Plugin {

    private static final int MAX_FILES = 20000;
    private JSObject pending;

    /* ------------------------------------------------------------ izin */

    private boolean hasAccess() {
        Context c = getContext();
        if (Build.VERSION.SDK_INT >= 30) {
            return Environment.isExternalStorageManager();
        }
        return c.checkSelfPermission(Manifest.permission.READ_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED
                && c.checkSelfPermission(Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED;
    }

    @PluginMethod
    public void checkAccess(PluginCall call) {
        JSObject r = new JSObject();
        r.put("granted", hasAccess());
        call.resolve(r);
    }

    @PluginMethod
    public void requestAccess(PluginCall call) {
        final Activity act = getActivity();
        act.runOnUiThread(() -> {
            try {
                if (Build.VERSION.SDK_INT >= 30) {
                    try {
                        Intent i = new Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION,
                                Uri.parse("package:" + getContext().getPackageName()));
                        act.startActivity(i);
                    } catch (ActivityNotFoundException e) {
                        act.startActivity(new Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION));
                    }
                } else {
                    act.requestPermissions(new String[]{
                            Manifest.permission.READ_EXTERNAL_STORAGE,
                            Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1001);
                }
                call.resolve();
            } catch (Exception e) {
                call.reject("Tidak bisa membuka pengaturan izin", e);
            }
        });
    }

    /* ------------------------------------------------------------ daftar berkas */

    @PluginMethod
    public void list(PluginCall call) {
        if (!hasAccess()) {
            call.reject("Tidak ada izin akses berkas", "no_access");
            return;
        }
        new Thread(() -> {
            try {
                File root = Environment.getExternalStorageDirectory();
                JSArray arr = new JSArray();
                ArrayDeque<File> stack = new ArrayDeque<>();
                stack.push(root);
                int count = 0;
                int dirs = 0;
                while (!stack.isEmpty() && count < MAX_FILES && dirs < 60000) {
                    File d = stack.pop();
                    dirs++;
                    File[] kids = d.listFiles();
                    if (kids == null) continue;
                    for (File f : kids) {
                        String n = f.getName();
                        if (n.startsWith(".")) continue;
                        if (f.isDirectory()) {
                            if (d.equals(root) && n.equals("Android")) continue;
                            stack.push(f);
                        } else {
                            if (n.startsWith("~$")) continue;
                            String ext = extOf(n);
                            if (!isDocExt(ext)) continue;
                            JSObject o = new JSObject();
                            o.put("path", f.getAbsolutePath());
                            o.put("name", n);
                            o.put("size", f.length());
                            o.put("mtime", f.lastModified());
                            o.put("ext", ext);
                            arr.put(o);
                            count++;
                            if (count >= MAX_FILES) break;
                        }
                    }
                }
                JSObject res = new JSObject();
                res.put("files", arr);
                res.put("count", count);
                call.resolve(res);
            } catch (Exception e) {
                call.reject("Gagal memindai berkas", e);
            }
        }).start();
    }

    private static String extOf(String name) {
        int i = name.lastIndexOf('.');
        return i >= 0 ? name.substring(i + 1).toLowerCase(Locale.ROOT) : "";
    }

    private static boolean isDocExt(String ext) {
        return ext.equals("docx") || ext.equals("xlsx") || ext.equals("pptx") || ext.equals("pdf");
    }

    /* ------------------------------------------------------------ baca / tulis */

    @PluginMethod
    public void readFile(PluginCall call) {
        final String uri = call.getString("uri", "");
        if (uri == null || uri.isEmpty()) {
            call.reject("Alamat berkas kosong");
            return;
        }
        new Thread(() -> {
            try {
                InputStream in;
                String name;
                if (uri.startsWith("content://") || uri.startsWith("file://")) {
                    Uri u = Uri.parse(uri);
                    in = getContext().getContentResolver().openInputStream(u);
                    name = queryName(u);
                } else {
                    File f = new File(uri);
                    in = new FileInputStream(f);
                    name = f.getName();
                }
                if (in == null) throw new Exception("Berkas tidak bisa dibuka");
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                try {
                    byte[] buf = new byte[64 * 1024];
                    int r;
                    while ((r = in.read(buf)) > 0) bos.write(buf, 0, r);
                } finally {
                    in.close();
                }
                byte[] bytes = bos.toByteArray();
                JSObject res = new JSObject();
                res.put("data", Base64.encodeToString(bytes, Base64.NO_WRAP));
                res.put("name", name);
                res.put("size", bytes.length);
                call.resolve(res);
            } catch (Exception e) {
                call.reject(e.getMessage() != null ? e.getMessage() : "Gagal membaca berkas", e);
            }
        }).start();
    }

    /** Menulis berkas: menimpa "path" bila diberikan, jika tidak membuat berkas baru di folder Download. */
    @PluginMethod
    public void writeFile(PluginCall call) {
        final String data = call.getString("data", "");
        String name = call.getString("name", "berkas");
        final String path = call.getString("path", "");
        if (data == null || data.isEmpty()) {
            call.reject("Data kosong");
            return;
        }
        final byte[] bytes;
        try {
            bytes = Base64.decode(data, Base64.DEFAULT);
        } catch (IllegalArgumentException e) {
            call.reject("Data tidak valid");
            return;
        }
        final String safeName = sanitize(name);
        new Thread(() -> {
            try {
                File target;
                if (path != null && !path.isEmpty()) {
                    target = new File(path);
                    File parent = target.getParentFile();
                    if (parent != null && !parent.exists()) parent.mkdirs();
                } else {
                    File dir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS);
                    if (!dir.exists()) dir.mkdirs();
                    target = uniqueFile(dir, safeName);
                }
                try (FileOutputStream os = new FileOutputStream(target)) {
                    os.write(bytes);
                }
                MediaScannerConnection.scanFile(getContext(), new String[]{target.getAbsolutePath()}, null, null);
                JSObject res = new JSObject();
                res.put("path", target.getAbsolutePath());
                res.put("name", target.getName());
                res.put("size", target.length());
                res.put("mtime", target.lastModified());
                call.resolve(res);
            } catch (Exception e) {
                call.reject(e.getMessage() != null ? e.getMessage() : "Gagal menyimpan", e);
            }
        }).start();
    }

    /* ------------------------------------------------------------ kelola */

    @PluginMethod
    public void rename(PluginCall call) {
        String path = call.getString("path", "");
        String newName = sanitize(call.getString("newName", ""));
        if (path == null || path.isEmpty() || newName.isEmpty()) {
            call.reject("Nama tidak valid");
            return;
        }
        File src = new File(path);
        File dst = new File(src.getParentFile(), newName);
        if (!src.exists()) {
            call.reject("Berkas tidak ditemukan");
            return;
        }
        if (dst.exists()) {
            call.reject("Sudah ada berkas dengan nama itu");
            return;
        }
        if (!src.renameTo(dst)) {
            call.reject("Gagal mengubah nama");
            return;
        }
        MediaScannerConnection.scanFile(getContext(), new String[]{src.getAbsolutePath(), dst.getAbsolutePath()}, null, null);
        JSObject r = new JSObject();
        r.put("path", dst.getAbsolutePath());
        r.put("name", dst.getName());
        call.resolve(r);
    }

    @PluginMethod
    public void delete(PluginCall call) {
        String path = call.getString("path", "");
        File f = new File(path == null ? "" : path);
        if (!f.exists()) {
            call.reject("Berkas tidak ditemukan");
            return;
        }
        if (!f.delete()) {
            call.reject("Gagal menghapus berkas");
            return;
        }
        MediaScannerConnection.scanFile(getContext(), new String[]{f.getAbsolutePath()}, null, null);
        call.resolve();
    }

    /* ------------------------------------------------------------ bagikan / buka dengan */

    private Uri uriFor(File f) {
        return FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", f);
    }

    @PluginMethod
    public void share(PluginCall call) {
        final String path = call.getString("path", "");
        final String mime = call.getString("mime", "*/*");
        final String pkg = call.getString("pkg", "");
        final File f = new File(path == null ? "" : path);
        if (!f.exists()) {
            call.reject("Berkas tidak ditemukan");
            return;
        }
        final Activity act = getActivity();
        act.runOnUiThread(() -> {
            try {
                Uri uri = uriFor(f);
                Intent i = new Intent(Intent.ACTION_SEND);
                i.setType(mime);
                i.putExtra(Intent.EXTRA_STREAM, uri);
                i.setClipData(ClipData.newRawUri(f.getName(), uri));
                i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                boolean fallback = false;
                if (pkg != null && !pkg.isEmpty()) {
                    Intent p = new Intent(i);
                    p.setPackage(pkg);
                    try {
                        act.startActivity(p);
                        JSObject r = new JSObject();
                        r.put("fallback", false);
                        call.resolve(r);
                        return;
                    } catch (ActivityNotFoundException e) {
                        fallback = true;
                    }
                }
                act.startActivity(Intent.createChooser(i, "Bagikan"));
                JSObject r = new JSObject();
                r.put("fallback", fallback);
                call.resolve(r);
            } catch (Exception e) {
                call.reject("Gagal membagikan berkas", e);
            }
        });
    }

    @PluginMethod
    public void openWith(PluginCall call) {
        final String path = call.getString("path", "");
        final String mime = call.getString("mime", "*/*");
        final File f = new File(path == null ? "" : path);
        if (!f.exists()) {
            call.reject("Berkas tidak ditemukan");
            return;
        }
        final Activity act = getActivity();
        act.runOnUiThread(() -> {
            try {
                Uri uri = uriFor(f);
                Intent i = new Intent(Intent.ACTION_VIEW);
                i.setDataAndType(uri, mime);
                i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                act.startActivity(Intent.createChooser(i, "Buka dengan"));
                call.resolve();
            } catch (Exception e) {
                call.reject("Tidak ada aplikasi untuk membuka berkas ini", e);
            }
        });
    }

    @PluginMethod
    public void pinShortcut(PluginCall call) {
        String path = call.getString("path", "");
        String name = call.getString("name", "Dokumen");
        JSObject r = new JSObject();
        try {
            if (Build.VERSION.SDK_INT < 26) {
                r.put("supported", false);
                call.resolve(r);
                return;
            }
            Context ctx = getContext();
            ShortcutManager sm = ctx.getSystemService(ShortcutManager.class);
            if (sm == null || !sm.isRequestPinShortcutSupported()) {
                r.put("supported", false);
                call.resolve(r);
                return;
            }
            Intent li = new Intent(ctx, MainActivity.class);
            li.setAction(Intent.ACTION_VIEW);
            li.putExtra("office_path", path);
            String label = name.length() > 24 ? name.substring(0, 24) : name;
            ShortcutInfo info = new ShortcutInfo.Builder(ctx, "doc-" + Integer.toHexString((path == null ? "" : path).hashCode()))
                    .setShortLabel(label)
                    .setLongLabel(name)
                    .setIcon(Icon.createWithResource(ctx, R.mipmap.ic_launcher))
                    .setIntent(li)
                    .build();
            sm.requestPinShortcut(info, null);
            r.put("supported", true);
            call.resolve(r);
        } catch (Exception e) {
            call.reject("Gagal membuat jalan pintas", e);
        }
    }

    /* ------------------------------------------------------------ dibuka dari aplikasi lain / jalan pintas */

    @Override
    public void load() {
        Activity a = getActivity();
        if (a != null) {
            pending = extractLaunch(a.getIntent());
            if (pending != null) a.setIntent(new Intent());
        }
    }

    @Override
    protected void handleOnNewIntent(Intent intent) {
        super.handleOnNewIntent(intent);
        JSObject p = extractLaunch(intent);
        if (p != null) {
            pending = p;
            notifyListeners("launch", new JSObject(), true);
        }
    }

    @PluginMethod
    public void getLaunch(PluginCall call) {
        JSObject r = pending != null ? pending : new JSObject();
        pending = null;
        call.resolve(r);
    }

    private JSObject extractLaunch(Intent i) {
        if (i == null || !Intent.ACTION_VIEW.equals(i.getAction())) return null;
        String p = i.getStringExtra("office_path");
        if (p != null && !p.isEmpty()) {
            JSObject r = new JSObject();
            r.put("path", p);
            r.put("name", new File(p).getName());
            return r;
        }
        Uri d = i.getData();
        if (d != null) {
            JSObject r = new JSObject();
            r.put("uri", d.toString());
            r.put("name", queryName(d));
            String t = i.getType();
            if (t != null) r.put("mime", t);
            return r;
        }
        return null;
    }

    /* ------------------------------------------------------------ util */

    private String queryName(Uri u) {
        String name = null;
        if ("content".equals(u.getScheme())) {
            ContentResolver cr = getContext().getContentResolver();
            try (Cursor c = cr.query(u, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
                if (c != null && c.moveToFirst()) name = c.getString(0);
            } catch (Exception ignored) {
            }
        }
        if (name == null || name.isEmpty()) name = u.getLastPathSegment();
        return name == null ? "berkas" : name;
    }

    private static String sanitize(String name) {
        if (name == null) return "";
        return name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
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
