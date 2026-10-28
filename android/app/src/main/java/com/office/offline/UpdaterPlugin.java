package com.office.offline;

import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageInfo;
import android.content.pm.PackageManager;
import android.content.pm.Signature;
import android.net.ConnectivityManager;
import android.net.Uri;
import android.os.Build;
import android.provider.Settings;

import androidx.core.content.FileProvider;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.MessageDigest;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Plugin "Updater": pembaruan aplikasi lewat GitHub Releases (sama dengan Alethra Browser).
 *
 * Workflow membuat rilis bertag "v1.N" (N = nomor run = versionCode) dan melampirkan APK, jadi
 * pembaruan = ada rilis dengan N lebih besar dari versionCode yang terpasang.
 *
 * Keamanan: APK hanya diunduh lewat HTTPS dari repositori yang ditanam saat build (alamat unduhan
 * diambil dari hasil check(), bukan dari JavaScript), diperiksa ukuran dan SHA-256 (bila GitHub
 * menyediakannya), dan tanda tangannya harus sama dengan aplikasi terpasang sebelum pemasangan dimulai.
 */
@CapacitorPlugin(name = "Updater")
public class UpdaterPlugin extends Plugin {

    private static final Pattern REPO = Pattern.compile("^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$");

    /** Rilis hasil check() terakhir; download() memakai ini, bukan data dari JS. */
    private static class Release {
        String tag, notes, url, sha256;
        long build, size;
    }

    private Release last;
    private final AtomicBoolean cancel = new AtomicBoolean(false);
    private volatile boolean busy = false;

    private String repo() {
        String r = BuildConfig.UPDATE_REPO == null ? "" : BuildConfig.UPDATE_REPO.trim();
        return REPO.matcher(r).matches() ? r : "";
    }

    private File dir() {
        return new File(getContext().getCacheDir(), "update");
    }

    private File apkFile(long build) {
        return new File(dir(), "Alethra-Office-" + build + ".apk");
    }

    @SuppressWarnings("deprecation")
    private long curBuild() {
        try {
            PackageInfo pi = getContext().getPackageManager().getPackageInfo(getContext().getPackageName(), 0);
            return Build.VERSION.SDK_INT >= 28 ? pi.getLongVersionCode() : pi.versionCode;
        } catch (Exception e) {
            return 0L;
        }
    }

    private String curVersion() {
        try {
            String v = getContext().getPackageManager().getPackageInfo(getContext().getPackageName(), 0).versionName;
            return v == null ? "?" : v;
        } catch (Exception e) {
            return "?";
        }
    }

    private boolean metered() {
        try {
            ConnectivityManager cm = (ConnectivityManager) getContext().getSystemService(Context.CONNECTIVITY_SERVICE);
            return cm != null && cm.isActiveNetworkMetered();
        } catch (Exception e) {
            return false;
        }
    }

    private boolean canInstall() {
        return Build.VERSION.SDK_INT < 26 || getContext().getPackageManager().canRequestPackageInstalls();
    }

    /* ------------------------------------------------------------ info */

    @PluginMethod
    public void info(PluginCall call) {
        JSObject r = new JSObject();
        r.put("enabled", !repo().isEmpty());
        r.put("repo", repo());
        r.put("version", curVersion());
        r.put("build", curBuild());
        r.put("metered", metered());
        call.resolve(r);
    }

    /* ------------------------------------------------------------ periksa */

    @PluginMethod
    public void check(PluginCall call) {
        if (repo().isEmpty()) {
            call.reject("Pembaruan otomatis tidak tersedia di build ini (alamat repositori GitHub tidak diketahui)");
            return;
        }
        new Thread(() -> {
            HttpURLConnection c = null;
            try {
                c = (HttpURLConnection) new URL("https://api.github.com/repos/" + repo() + "/releases/latest").openConnection();
                c.setConnectTimeout(15000);
                c.setReadTimeout(20000);
                c.setRequestProperty("User-Agent", "AlethraOffice");
                c.setRequestProperty("Accept", "application/vnd.github+json");
                int code = c.getResponseCode();
                JSObject res = new JSObject();
                res.put("build", curBuild());
                res.put("version", curVersion());
                if (code == 404) {
                    last = null;
                    res.put("available", false);
                    res.put("norelease", true);
                    call.resolve(res);
                    return;
                }
                if (code == 403 || code == 429) throw new RuntimeException("Batas permintaan GitHub tercapai, coba lagi nanti");
                if (code != 200) throw new RuntimeException("GitHub menjawab " + code);
                JSONObject js = new JSONObject(readAll(c.getInputStream()));
                String tag = js.optString("tag_name");
                Matcher m = Pattern.compile("(\\d+)$").matcher(tag);
                JSONArray assets = js.optJSONArray("assets");
                Release rel = null;
                if (m.find() && assets != null) {
                    for (int i = 0; i < assets.length(); i++) {
                        JSONObject a = assets.getJSONObject(i);
                        String url = a.optString("browser_download_url");
                        if (!a.optString("name").toLowerCase(Locale.ROOT).endsWith(".apk") || !url.startsWith("https://")) continue;
                        String dg = a.optString("digest");
                        rel = new Release();
                        rel.tag = tag;
                        rel.build = Long.parseLong(m.group(1));
                        rel.notes = js.optString("body");
                        rel.url = url;
                        rel.size = a.optLong("size", 0L);
                        rel.sha256 = dg.toLowerCase(Locale.ROOT).startsWith("sha256:") ? dg.substring(7).toLowerCase(Locale.ROOT) : null;
                        break;
                    }
                }
                last = rel;
                if (rel == null || rel.build <= curBuild()) {
                    res.put("available", false);
                } else {
                    res.put("available", true);
                    res.put("tag", rel.tag);
                    res.put("latest", rel.tag.startsWith("v") ? rel.tag.substring(1) : rel.tag);
                    res.put("latestBuild", rel.build);
                    res.put("notes", rel.notes == null ? "" : rel.notes.trim());
                    res.put("size", rel.size);
                    res.put("downloaded", intact(apkFile(rel.build), rel));
                }
                call.resolve(res);
            } catch (Exception e) {
                call.reject(e.getMessage() != null ? e.getMessage() : "Gagal memeriksa");
            } finally {
                if (c != null) c.disconnect();
            }
        }).start();
    }

    /* ------------------------------------------------------------ unduh */

    @PluginMethod
    public void download(PluginCall call) {
        final Release r = last;
        if (r == null) {
            call.reject("Periksa pembaruan dulu");
            return;
        }
        if (busy) {
            call.reject("Unduhan sedang berjalan");
            return;
        }
        busy = true;
        cancel.set(false);
        new Thread(() -> {
            File part = null;
            HttpURLConnection c = null;
            try {
                File d = dir();
                d.mkdirs();
                File out = apkFile(r.build);
                if (intact(out, r)) {
                    JSObject res = new JSObject();
                    res.put("ok", true);
                    call.resolve(res);
                    return;
                }
                out.delete();
                File[] old = d.listFiles();
                if (old != null) for (File f : old) if (!f.equals(out)) f.delete();   // sisa unduhan versi lain
                part = new File(d, out.getName() + ".part");
                part.delete();
                c = (HttpURLConnection) new URL(r.url).openConnection();
                c.setConnectTimeout(15000);
                c.setReadTimeout(30000);
                c.setInstanceFollowRedirects(true);
                c.setRequestProperty("User-Agent", "AlethraOffice");
                if (c.getResponseCode() != 200) throw new RuntimeException("Server menjawab " + c.getResponseCode());
                long total = r.size > 0 ? r.size : c.getContentLengthLong();
                long got = 0, lastT = 0;
                try (FileOutputStream o = new FileOutputStream(part); InputStream in = c.getInputStream()) {
                    byte[] buf = new byte[64 * 1024];
                    while (true) {
                        if (cancel.get()) {
                            part.delete();
                            JSObject res = new JSObject();
                            res.put("cancelled", true);
                            call.resolve(res);
                            return;
                        }
                        int n = in.read(buf);
                        if (n < 0) break;
                        o.write(buf, 0, n);
                        got += n;
                        long now = System.currentTimeMillis();
                        if (now - lastT > 250) {
                            lastT = now;
                            progress(got, total);
                        }
                    }
                }
                progress(got, total);
                if (!part.renameTo(out)) {
                    part.delete();
                    throw new RuntimeException("Gagal menyimpan berkas");
                }
                if (!intact(out, r)) {
                    out.delete();
                    throw new RuntimeException("Berkas unduhan rusak (ukuran atau checksum tidak cocok)");
                }
                JSObject res = new JSObject();
                res.put("ok", true);
                call.resolve(res);
            } catch (Exception e) {
                if (part != null) part.delete();
                call.reject(e.getMessage() != null ? e.getMessage() : "Gagal mengunduh");
            } finally {
                busy = false;
                if (c != null) c.disconnect();
            }
        }).start();
    }

    private void progress(long done, long total) {
        JSObject p = new JSObject();
        p.put("done", done);
        p.put("total", total);
        notifyListeners("progress", p);
    }

    @PluginMethod
    public void cancel(PluginCall call) {
        cancel.set(true);
        call.resolve();
    }

    /* ------------------------------------------------------------ pasang */

    @PluginMethod
    public void install(PluginCall call) {
        final Release r = last;
        if (r == null) {
            call.reject("Periksa pembaruan dulu");
            return;
        }
        final File f = apkFile(r.build);
        if (!f.isFile()) {
            call.reject("Berkas pembaruan belum diunduh");
            return;
        }
        JSObject res = new JSObject();
        Boolean same = sameSigner(f);
        if (Boolean.FALSE.equals(same)) {
            f.delete();
            res.put("status", "badsig");
            call.resolve(res);
            return;
        }
        if (!canInstall()) {
            res.put("status", "needperm");
            call.resolve(res);
            return;
        }
        try {
            Uri uri = FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", f);
            Intent i = new Intent(Intent.ACTION_VIEW)
                    .setDataAndType(uri, "application/vnd.android.package-archive")
                    .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(i);
            res.put("status", "started");
            call.resolve(res);
        } catch (Exception e) {
            call.reject("Tidak bisa memulai pemasangan");
        }
    }

    @PluginMethod
    public void openInstallSettings(PluginCall call) {
        try {
            Intent i = new Intent(Settings.ACTION_MANAGE_UNKNOWN_APP_SOURCES, Uri.parse("package:" + getContext().getPackageName()))
                    .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(i);
            call.resolve();
        } catch (Exception e) {
            call.reject("Pengaturan tidak tersedia");
        }
    }

    /** Hapus APK sisa pembaruan yang sudah terpasang atau kedaluwarsa (build <= versi terpasang). */
    @PluginMethod
    public void cleanup(PluginCall call) {
        final long installed = curBuild();
        new Thread(() -> {
            try {
                File[] fs = dir().listFiles();
                if (fs != null) {
                    Pattern p = Pattern.compile("Alethra-Office-(\\d+)\\.apk");
                    for (File f : fs) {
                        Matcher m = p.matcher(f.getName());
                        if (!m.matches() || Long.parseLong(m.group(1)) <= installed) f.delete();
                    }
                }
            } catch (Exception ignored) {
            }
            call.resolve();
        }).start();
    }

    /* ------------------------------------------------------------ util */

    private static String readAll(InputStream in) throws Exception {
        try (InputStream is = in; java.io.ByteArrayOutputStream bo = new java.io.ByteArrayOutputStream()) {
            byte[] b = new byte[8192];
            int n;
            while ((n = is.read(b)) > 0) bo.write(b, 0, n);
            return bo.toString("UTF-8");
        }
    }

    private static String sha256(File f) throws Exception {
        MessageDigest md = MessageDigest.getInstance("SHA-256");
        try (InputStream in = new java.io.FileInputStream(f)) {
            byte[] b = new byte[64 * 1024];
            int n;
            while ((n = in.read(b)) > 0) md.update(b, 0, n);
        }
        return hex(md.digest());
    }

    private static String hex(byte[] d) {
        StringBuilder sb = new StringBuilder();
        for (byte x : d) sb.append(String.format("%02x", x));
        return sb.toString();
    }

    private boolean intact(File f, Release r) {
        try {
            return f.isFile() && (r.size <= 0 || f.length() == r.size) && (r.sha256 == null || sha256(f).equals(r.sha256));
        } catch (Exception e) {
            return false;
        }
    }

    private static Set<String> certHashes(Signature[] sigs) throws Exception {
        Set<String> s = new HashSet<>();
        if (sigs != null) for (Signature g : sigs) s.add(hex(MessageDigest.getInstance("SHA-256").digest(g.toByteArray())));
        return s;
    }

    /** true = APK ditandatangani kunci yang sama dengan aplikasi terpasang, false = beda, null = tidak bisa dipastikan. */
    @SuppressWarnings("deprecation")
    private Boolean sameSigner(File apk) {
        try {
            PackageManager pm = getContext().getPackageManager();
            Set<String> mine, theirs;
            if (Build.VERSION.SDK_INT >= 28) {
                int fl = PackageManager.GET_SIGNING_CERTIFICATES;
                PackageInfo a = pm.getPackageInfo(getContext().getPackageName(), fl);
                PackageInfo b = pm.getPackageArchiveInfo(apk.getPath(), fl);
                mine = certHashes(a.signingInfo == null ? null : a.signingInfo.getApkContentsSigners());
                theirs = b == null || b.signingInfo == null ? new HashSet<>() : certHashes(b.signingInfo.getApkContentsSigners());
            } else {
                int fl = PackageManager.GET_SIGNATURES;
                PackageInfo a = pm.getPackageInfo(getContext().getPackageName(), fl);
                PackageInfo b = pm.getPackageArchiveInfo(apk.getPath(), fl);
                mine = certHashes(a.signatures);
                theirs = b == null ? new HashSet<>() : certHashes(b.signatures);
            }
            if (mine.isEmpty() || theirs.isEmpty()) return null;
            return mine.equals(theirs);
        } catch (Exception e) {
            return null;
        }
    }
}
