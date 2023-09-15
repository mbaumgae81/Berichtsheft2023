package dev.baumgaertner;

import java.io.*;
import java.nio.channels.FileChannel;

public class MyCopy {
    public static boolean copyFile(final File src, final File dest) throws IOException, FileNotFoundException {
        FileChannel srcChannel = new FileInputStream(src).getChannel();
        FileChannel destChannel = new FileOutputStream(dest).getChannel();
        try {
            srcChannel.transferTo(0, srcChannel.size(), destChannel);
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        } finally {
            if (srcChannel != null)
                srcChannel.close();
            if (destChannel != null)
                destChannel.close();
        }
        return true;
    }
}
