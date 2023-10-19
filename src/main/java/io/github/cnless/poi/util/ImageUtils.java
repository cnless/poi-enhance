package io.github.cnless.poi.util;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.Document;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.util.Objects;


public class ImageUtils {
    public static byte[] toByteArray(InputStream is) {
        if (null == is) return null;
        try {
            return IOUtils.toByteArray(is);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(is);
        }
        return null;
    }

    public static InputStream getUrlStream(String urlPath) throws IOException {
        URL url = new URL(urlPath);
        URLConnection connection = url.openConnection();
        connection.addRequestProperty("User-Agent", "Mozilla/4.0");
        InputStream inputStream = connection.getInputStream();
        if (connection instanceof HttpURLConnection) {
            if (200 != ((HttpURLConnection) connection).getResponseCode()) {
                throw new IOException("get url " + urlPath + " content error, response status: "
                        + ((HttpURLConnection) connection).getResponseCode());
            }
        }
        return inputStream;
    }

    public static byte[] toByteArray(String urlPath){
        InputStream inputStream = null;
        try {
            inputStream = getUrlStream(urlPath);
            return IOUtils.toByteArray(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (Objects.nonNull(inputStream)){
                IOUtils.closeQuietly(inputStream);
            }
        }
        return null;
    }

    public static byte[] getLocalByteArray(File res) {
        try {
            return Files.readAllBytes(res.toPath());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static PictureType suggestFileType(byte[] bytes) {
        if (startsWith(bytes, "GIF89a".getBytes()) || startsWith(bytes, "GIF87a".getBytes())) {
            return PictureType.GIF;
        }
        if (startsWith(bytes, new byte[] { (byte) 0xFF, (byte) 0xD8 })
                || endsWith(bytes, new byte[] { (byte) 0xFF, (byte) 0xD9 })) {
            return PictureType.JPEG;
        }
        if (startsWith(bytes, new byte[] { (byte) 0x89, (byte) 0x50, (byte) 0x4E, (byte) 0x47 })) {
            return PictureType.PNG;
        }
        if (startsWith(bytes, new byte[] { (byte) 0x49, (byte) 0x49, (byte) 0x2A, (byte) 0x00 })
                || startsWith(bytes, new byte[] { (byte) 0x4D, (byte) 0x4D, (byte) 0x00, (byte) 0x2A })) {
            return PictureType.TIFF;
        }
        if (startsWith(bytes, "BM".getBytes())) {
            return PictureType.BMP;
        }
        throw new IllegalArgumentException("Unable to identify the picture type from byte");
    }

    public static boolean startsWith(byte[] bytes, byte[] prefix) {
        if (bytes == prefix) return true;
        if (null == prefix || null == bytes || bytes.length < prefix.length) return false;
        for (int i = 0; i < prefix.length; i++) {
            if (bytes[i] != prefix[i]) return false;
        }
        return true;
    }

    public static boolean endsWith(byte[] bytes, byte[] suffix) {
        if (bytes == suffix) return true;
        if (null == suffix || null == bytes || bytes.length < suffix.length) return false;
        int length = bytes.length - suffix.length;
        for (int i = suffix.length - 1; i >= 0; i--) {
            if (bytes[length + i] != suffix[i]) return false;
        }
        return true;
    }

}
