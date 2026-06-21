package com.example.excelparesernew;

import java.io.*;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class FileUtils {

    static void unzip(Path source, Path destDir, int bufferSize) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(
                new BufferedInputStream(Files.newInputStream(source), bufferSize))) {
            ZipEntry entry;
            byte[] buf = new byte[bufferSize];
            while ((entry = zis.getNextEntry()) != null) {
                Path target = destDir.resolve(entry.getName());
                if (entry.isDirectory()) {
                    Files.createDirectories(target);
                } else {
                    Files.createDirectories(target.getParent());
                    try (FileOutputStream fos = new FileOutputStream(target.toFile())) {
                        int len;
                        while ((len = zis.read(buf)) > 0) fos.write(buf, 0, len);
                    }
                }
            }
        }
    }

    static void zip(Path sourceDir, Path outputPath, int bufferSize) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(
                new BufferedOutputStream(Files.newOutputStream(outputPath), bufferSize));
             Stream<Path> stream = Files.walk(sourceDir)) {
            stream.forEach(path -> {
                try {
                    String rel = sourceDir.relativize(path).toString().replace("\\", "/");
                    if (Files.isDirectory(path)) {
                        if (!rel.isEmpty()) {
                            zos.putNextEntry(new ZipEntry(rel + "/"));
                            zos.closeEntry();
                        }
                    } else {
                        zos.putNextEntry(new ZipEntry(rel));
                        Files.copy(path, zos);
                        zos.closeEntry();
                    }
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            });
        }
    }

    static void deleteDirectory(Path dir) throws IOException {
        if (Files.exists(dir)) {
            Files.walkFileTree(dir, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Files.delete(file);          // удаляем файл
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                    Files.delete(dir);           // после обхода содержимого удаляем саму директорию
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult visitFileFailed(Path file, IOException exc) throws IOException {
                    // при ошибке доступа к файлу — бросаем исключение
                    throw exc;
                }
            });
        }
    }

}
