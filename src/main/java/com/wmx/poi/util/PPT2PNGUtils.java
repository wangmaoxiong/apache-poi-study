/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package com.wmx.poi.util;

import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.util.PPTX2PNG;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.lang.ref.WeakReference;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.TreeSet;
import java.util.logging.Logger;

/**
 * Demonstrates how you can use HSLF to convert each slide into a PNG image
 */
public class PPT2PNGUtils extends PPTX2PNG {
    private static Logger logger = Logger.getAnonymousLogger();

    public static void main(String[] args) {
        File file = new File("C:\\Users\\22684\\Desktop\\apachecon_eu_08.pptx");
        ppt2png(file);
    }

    /**
     * ppt文件(.ppt、.pptx)转换为图片
     *
     * @param pptFile   待转换的 ppt 文件。注意如果生成的图片出现中文乱码，则可以将 ppt 文件中的中文设置为 "宋体"，亲测转换有效
     * @param outPngDir 转换后，图片的存放目录，默认为 pptFile 文件同一级目录，并使用 ppt 文件名新建一个目录名
     * @param format    转换后的图片格式，支持 jpg、jpeg、png、gif，默认为 png
     * @param range     ppt转换的页面范围，如 1,3,5-8,9,10 表示转换第 1，3，5页，以及第5至8页，和第 9，10页.默认转换所有
     * @param scale     幻灯片转图片的比例，如 1.0 表示幻灯片与图片进行 1:1 转换；0.5 表示图片的尺寸是幻灯片尺寸的一半
     * @return
     */

    public static void ppt2png(File pptFile, File outPngDir, String format, String range, Float scale) {
        if (pptFile == null || !pptFile.exists() || pptFile.isDirectory()) {
            logger.info(pptFile + " 是非法的 ppt 文件...");
            return;
        }
        if (outPngDir == null || !outPngDir.exists() || outPngDir.isFile()) {
            outPngDir = new File(pptFile.getParentFile(), pptFile.getName().substring(0, pptFile.getName().lastIndexOf(".")));
            outPngDir.mkdirs();
        }
        format = (format == null || !"jpg jpeg png gif".contains(format.toLowerCase())) ? "png" : format;
        scale = (scale == null || scale < 0) ? 1.0f : scale;

        /**
         * SlideShow<S,P> create(File file, String password, boolean readOnly)
         * 从给定的 ppt 文件创建适当的 Hslfslideshow 与 Xmlsideshow，文件必须存在且可读
         * 文件受密码保护时，需要提供密码，否则没有密码时，设置为 null 即可
         * 以只读模式打开幻灯片可以避免回写。为了正确释放资源，使用后应关闭幻灯片放映。
         */
        SlideShow<?, ?> slideShow = null;
        try {
            slideShow = SlideShowFactory.create(pptFile, null, true);
        } catch (IOException e) {
            e.printStackTrace();
        }
        List<? extends Slide<?, ?>> slideList = slideShow.getSlides();//获取所有幻灯片
        Set<Integer> slideIndexSet = convertedPageIndex(slideList.size(), range);//获取需要转换的幻灯片索引集合
        logger.info("转换 " + pptFile.getAbsolutePath() + " ppt 的幻灯片索引为：" + slideIndexSet);
        if (slideIndexSet.isEmpty()) {
            logger.info(pptFile + " 文件内容为空，或者范围无效：" + range + ", 放弃转换。");
            return;
        }
        Dimension pgDimension = slideShow.getPageSize();//获取幻灯片尺寸
        int width = Math.round(pgDimension.width * scale);
        int height = Math.round(pgDimension.height * scale);
        for (Integer slideNo : slideIndexSet) {
            logger.info("开始转换第 " + (slideNo + 1) + " 张幻灯片...");
            Slide<?, ?> slide = slideList.get(slideNo);
            BufferedImage bufferedImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D graphics = bufferedImage.createGraphics();

            // default rendering options
            graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
            graphics.setRenderingHint(Drawable.BUFFERED_IMAGE, new WeakReference<>(bufferedImage));

            graphics.scale(scale, scale);
            try {
                slide.draw(graphics);//这个方法容易抛异常，需要捕获处理,让它继续转换
            } catch (Exception e) {
                logger.info("第 " + (slideNo + 1) + " 张幻灯片转换失败...");
                //continue;//如果继续往后执行，则会创建一张空白的图片。如果不需要空白图片，则可以直接进入下一张转换
            }
            String outName = pptFile.getName().replaceFirst(".pptx?", "");
            outName = String.format(Locale.ROOT, "%1$s-%2$04d.%3$s", outName, slideNo + 1, format);
            File outfile = new File(outPngDir, outName);
            try {
                ImageIO.write(bufferedImage, format, outfile);//文件已经存在时，会自动覆盖
            } catch (IOException e) {
                e.printStackTrace();
                logger.info("第 " + (slideNo + 1) + " 张幻灯片转换失败...");
                continue;
            }
            graphics.dispose();
            bufferedImage.flush();
            logger.info("第 " + (slideNo + 1) + " 张幻灯片转换完成...");
        }

        logger.info(pptFile + " 转换完成,文件输出目录：" + outPngDir.getAbsolutePath());
    }

    public static void ppt2png(File pptFile, File outPngDir, String format, String range) {
        ppt2png(pptFile, outPngDir, format, range, 1.0F);
    }

    public static void ppt2png(File pptFile, File outPngDir, String format) {
        ppt2png(pptFile, outPngDir, format, null, 1.0F);
    }

    public static void ppt2png(File pptFile, File outPngDir) {
        ppt2png(pptFile, outPngDir, null, null, 1.0F);
    }

    public static void ppt2png(File pptFile) {
        ppt2png(pptFile, null, null, null, 1.0F);
    }

    /**
     * 整理 ppt 幻灯片转换的索引集合
     *
     * @param slideCount ppt 幻灯片总数
     * @param range      需要转换为图片的幻灯片索引范围，如 1,3,5-8,9,10 表示转换第 1，3，5页，以及第5至8页，和第 9，10页.默认转换所有
     * @return
     */
    private static Set<Integer> convertedPageIndex(final int slideCount, String range) {
        Set<Integer> slideIdx = new TreeSet<>();//使用 set 自动去重
        if (range == null || "".equals(range.trim())) {
            for (int i = 0; i < slideCount; i++) {
                slideIdx.add(i);
            }
            return slideIdx;
        }
        for (String subrange : range.split(",")) {
            String[] idx = subrange.split("-");
            switch (idx.length) {
                case 0:
                    break;
                case 1:
                    int subIdx = Math.max(Integer.parseInt(idx[0]), 1);
                    slideIdx.add(Math.min(subIdx, slideCount) - 1);
                    break;
                case 2:
                    idx[0] = idx[0].equals("") ? "1" : idx[0];//防止用户输入非法的负数
                    int startIdx = Math.min(Integer.parseInt(idx[0]), slideCount);
                    int endIdx = Math.min(Integer.parseInt(idx[1]), slideCount);
                    for (int i = Math.max(startIdx, 1); i <= endIdx; i++) {
                        slideIdx.add(i - 1);
                    }
                    break;
                default:
            }
        }
        return slideIdx;
    }
}
