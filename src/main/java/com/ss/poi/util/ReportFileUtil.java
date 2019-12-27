package com.ss.poi.util;

import com.ss.poi.entity.SensorPOIException;
import com.ss.poi.entity.TestResultData;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import java.io.*;

/**
 * @Description
 */
@SuppressWarnings("restriction")
public class ReportFileUtil {

	public static void checkFilePathExists(String fileLocation) {

		File file = new File(fileLocation);
		// 获取文件的父路径字符串
		File path = file.getParentFile();
		if (!path.exists()) {
			boolean isCreated = path.mkdirs();
			if (!isCreated) {
			}
		}
	}

	public static boolean checkFileLocationExists(String fileLocation) {

		File file = new File(fileLocation);
		checkFilePathExists(fileLocation);
		return file.exists();
	}

	/**
	 * 创建word模板,若文件不存在则创建新模板
	 * 
	 * 
	 */
	public static XWPFDocument openDocx(String fileLocation) throws Exception {

		XWPFDocument doc;
		if (fileLocation == null || fileLocation == "") {
			doc = new XWPFDocument();
		} else {
			String fileName = fileLocation.substring(fileLocation
					.lastIndexOf("/") + 1);
			if (fileName.indexOf(".") == -1) {
				fileLocation += ".docx";
			}

			File file = new File(fileLocation);
			InputStream is;
			if (!file.exists()) {
				// is = new FileInputStream(getReportTemplete());
				doc = new XWPFDocument();
				return doc;
			} else {
				// OPCPackage opcPackage = POIXMLDocument
				// .openPackage(fileLocation);
				is = new FileInputStream(fileLocation);
			}
			doc = new XWPFDocument(is);
		}
		return doc;
	}

	public static String getLiuChengTempleteLocation(TestResultData data) {
		return "./Report/Templete/" + data.getTestLiuCheng() + ".docx";
	}

	public static String getFileLocationByData(TestResultData data) {
		String fileLocation = "./Report/" + data.getTestedID() + "/"
				+ data.getMissionID() + "/" + data.getDataIndex2() + "/"
				+ data.getDateStr();
		return fileLocation;
	}

	public static String getFileLocationByMissionID(int pdid) {
		String fileLocation = "./Report/Final/" + pdid + ".docx";
		return fileLocation;
	}

	public static String getFileLocationByIndexDate(int pdid, String testedID,
			String index, String date) {
		String fileLocation = "./Report/" + testedID + "/" + pdid + "/" + index
				+ "/" + date + ".docx";
		return fileLocation;
	}

	public static String getFilePathByIndex(int pdid, String testedID,
			String index) {
		String fileLocation = "./Report/" + testedID + "/" + pdid + "/" + index;
		return fileLocation;
	}

	public static String getTestResultFileLocation(TestResultData data) {
		return "./TestResult/" + data.getTestedID() + "/" + data.missionID
				+ "/" + data.getDataIndex2() + "/" + data.dateStr + ".json";
	}

	private static String getReportTemplete() {
		return "./Report/Templete.docx";
	}

	public static String getHtmlUrlByMissionID(String testedID, int pdid,
			String time) {
		String fileLocation = "./Report/" + testedID + "/" + pdid + "/" + time
				+ ".html";
		return fileLocation;
	}

	private static void checkFileLocation(String fileLocation)
			throws SensorPOIException {
		if (fileLocation == null || fileLocation == "") {
			throw new SensorPOIException(-1);
		}
		File file = new File(fileLocation);
		// 获取文件的父路径字符串
		File path = file.getParentFile();
		if (!path.exists()) {
			System.out.println("文件夹不存在，自动创建：" + path);
			boolean isCreated = path.mkdirs();
			if (!isCreated) {
				System.out.println("创建文件夹失败，path=" + path);
				throw new SensorPOIException(-1);
			}
		}
	}

	/**
	 * 输出docx文件
	 * 
	 * @param doc
	 * @param fileName
	 *            文件名(不包含后缀)
	 * @param fileStyle
	 *            文件类型 doc、html
	 * @throws Exception
	 */
	public static void saveAsDocx(XWPFDocument doc, String fileLocation)
			throws SensorPOIException, Exception {
		// String fileStyle = filePath.substring(filePath.lastIndexOf(".") + 1);
		String fileName = fileLocation
				.substring(fileLocation.lastIndexOf("/") + 1);
		if (fileName.indexOf(".") == -1) {
			fileLocation += ".docx";
		}
		checkFileLocation(fileLocation);
		OutputStream os = new FileOutputStream(fileLocation);
		doc.write(os);
		os.close();
		System.out.println("docx已生成：" + fileLocation);
	}

	/**
	 * 输出docx文件
	 * 
	 * @param fileName
	 *            文件名(不包含后缀)
	 * @param fileStyle
	 *            文件类型 doc、html
	 * @throws Exception
	 */
	public static void saveAsHtml(XWPFDocument doc, String fileLocation)
			throws SensorPOIException, Exception {
		// String fileStyle = filePath.substring(filePath.lastIndexOf(".") + 1);
		File imageFolderFile = new File(fileLocation);

		String fileName = fileLocation
				.substring(fileLocation.lastIndexOf("/") + 1);
		if (fileName.indexOf(".") == -1) {
			fileLocation += ".html";
		}
		checkFileLocation(fileLocation);

		/*
		 * OutputStreamWriter outputStreamWriter = null; XHTMLOptions options =
		 * XHTMLOptions.create(); // 存放图片的文件夹 options.setExtractor(new
		 * FileImageExtractor(imageFolderFile)); // html中图片的路径
		 * options.URIResolver(new BasicURIResolver("image"));
		 * outputStreamWriter = new OutputStreamWriter(new
		 * FileOutputStream(fileLocation), "utf-8"); XHTMLConverter
		 * xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
		 * xhtmlConverter.convert(doc, outputStreamWriter, options);
		 */

		// 解析 XHTML配置 (这里设置IURIResolver来设置图片存放的目录)
		XHTMLOptions options = XHTMLOptions.create().URIResolver(
				new FileURIResolver(imageFolderFile));
		options.setExtractor(new FileImageExtractor(imageFolderFile));
		options.setIgnoreStylesIfUnused(false);
		options.setFragment(true);
		// )将 XWPFDocument转换成XHTML
		OutputStream out = new FileOutputStream(new File(fileLocation));
		XHTMLConverter.getInstance().convert(doc, out, options);
		out.close();
		System.out.println("html已生成：" + fileLocation);
	}

	/**
	 * 输出文件
	 * 
	 * @param doc
	 * @param fileName
	 *            文件名(不包含后缀)
	 * @param fileStyle
	 *            文件类型 doc、html
	 * @throws Exception
	 */
	public static void saveFile(XWPFDocument doc, String fileName,
			String fileStyle) throws Exception {
		// String fileStyle = filePath.substring(filePath.lastIndexOf(".") + 1);
		// String fileName = filePath.substring(filePath.lastIndexOf("/")+1);
		String outputPath = "e:/";
		String filePath = outputPath + fileName + "." + fileStyle;
		if (fileStyle.toLowerCase().contains("doc".toLowerCase())) {
			OutputStream os = new FileOutputStream(filePath);
			doc.write(os);
			if (os != null) {
				try {
					os.close();
					System.out.println("Word文件已输: " + filePath);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} else if (fileStyle.toLowerCase().contains("html".toLowerCase())) {

		}
	}

	public static void copyFile(String src, String dest) throws Exception {
		File file = new File(src);
		FileInputStream inputStream = new FileInputStream(file);
		FileOutputStream outputStream = new FileOutputStream(dest);
		byte[] buffer = new byte[(int) file.length()];
		inputStream.read(buffer);
		outputStream.write(buffer);
		outputStream.flush();
		outputStream.close();
		inputStream.close();
		System.out.println("复制文件: " + src + ",到: " + dest + "  成功!");
	}

	/**
	 * 将文件转成base64 字符串
	 * 
	 * @param path文件路径
	 * @return *
	 * @throws Exception
	 */
	public static String encodeBase64File(String filePath) {
		if (filePath == null || filePath == "" || filePath.equals("")) {
			System.out.println(filePath + "(下载路径不正确！)");
		}
		File file = new File(filePath);
		byte[] buffer = null;
		try {
			FileInputStream inputFile = new FileInputStream(file);
			buffer = new byte[(int) file.length()];
			inputFile.read(buffer);
			inputFile.close();
		} catch (IOException e) {
			System.out.println("下载文件不存在：" + filePath);
			return null;
		}
		return new BASE64Encoder().encode(buffer);
	}

	public static String TEMPIMGPATH = "./Report/image/";
	public static String getTempImageLocation(){
		return TEMPIMGPATH + System.currentTimeMillis() + ".jpg";
	}

	public static String byte64ToTempImg(String base64Str) throws IOException {
		BASE64Decoder decoder = new BASE64Decoder();
		byte[] decodeBuffer = decoder.decodeBuffer(base64Str);
		BufferedInputStream bis = null;
		FileOutputStream fos = null;
		BufferedOutputStream output = null;
		String fileLocation=getTempImageLocation();
		try {
			ByteArrayInputStream byteInputStream = new ByteArrayInputStream(
					decodeBuffer);
			bis = new BufferedInputStream(byteInputStream);

			File file = new File(fileLocation);
			// 获取文件的父路径字符串
			File path = file.getParentFile();
			if (!path.exists()) {
				System.out.println("文件夹不存在，自动创建：" + path);
				boolean isCreated = path.mkdirs();
				if (!isCreated) {
					System.out.println("创建文件夹失败，path=" + path);
				}
			}
			fos = new FileOutputStream(file);
			// 实例化OutputString 对象
			output = new BufferedOutputStream(fos);
			byte[] buffer = new byte[1024];
			int length = bis.read(buffer);
			while (length != -1) {
				output.write(buffer, 0, length);
				length = bis.read(buffer);
			}
			output.flush();
		} catch (Exception e) {
			System.out.println("输出文件流时抛异常，filePath=" + fileLocation);
		} finally {
			try {
				bis.close();
				fos.close();
				output.close();
			} catch (IOException e0) {
				System.out.println("文件处理失败，filePath=" + fileLocation);
			}
		}
		System.out.println("base64已生成文件：" + fileLocation);
		return fileLocation;
	}

	public static void delTempFile(String path) {
		if (!new File(path).delete()) {
			System.out.println("文件删除失败：" + path);
		}
		System.out.println("文件已删除：" + path);
	}

}
