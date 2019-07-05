package com.example.code.docx4j;

import com.example.code.excel.ImportExcelUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTAltChunk;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.Color;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Lv Jie
 * @version 1.0.0
 * @desc TODO
 * @create 2019/7/4 15:37
 */
public class Docx4jWordUtil {

	public static WordprocessingMLPackage wordMLPackage;

	static {
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}

	public static ObjectFactory factory = Context.getWmlObjectFactory();

	/**
	 * 本方法创建单元格, 添加样式后添加到表格行中
	 */
	public static Tr addStyledTableCell(Tr tableRow, String content,
                                        boolean bold, String fontSize) {
		Tc tableCell = factory.createTc();
		addStyling(tableCell, content, bold, fontSize);
		tableRow.getContent().add(tableCell);
		return tableRow;
	}

	/**
	 * 这里我们添加实际的样式信息, 首先创建一个段落, 然后创建以单元格内容作为值的文本对象;
	 * 第三步, 创建一个被称为运行块的对象, 它是一块或多块拥有共同属性的文本的容器, 并将文本对象添加
	 * 到其中. 随后我们将运行块R添加到段落内容中.
	 * 直到现在我们所做的还没有添加任何样式, 为了达到目标, 我们创建运行块属性对象并给它添加各种样式.
	 * 这些运行块的属性随后被添加到运行块. 最后段落被添加到表格的单元格中.
	 */
	private static Tc addStyling(Tc tableCell, String content, boolean bold, String fontSize) {
		P paragraph = factory.createP();
		Text text = factory.createText();
		text.setValue(content);
		R run = factory.createR();
		run.getContent().add(text);
		paragraph.getContent().add(run);
		RPr runProperties = factory.createRPr();
		if (bold) {
			addBoldStyle(runProperties);
		}
		if (fontSize != null && !fontSize.isEmpty()) {
			setFontSize(runProperties, fontSize);
		}
		run.setRPr(runProperties);
		tableCell.getContent().add(paragraph);
		return tableCell;
	}

	/**
	 * 本方法为可运行块添加字体大小信息. 首先创建一个"半点"尺码对象, 然后设置fontSize
	 * 参数作为该对象的值, 最后我们分别设置sz和szCs的字体大小.
	 * Finally we'll set the non-complex and complex script font sizes, sz and szCs respectively.
	 */
	public static void setFontSize(RPr runProperties, String fontSize) {
		HpsMeasure size = new HpsMeasure();
		size.setVal(new BigInteger(fontSize));
		runProperties.setSz(size);
		runProperties.setSzCs(size);
	}

	/**
	 * 本方法给可运行块属性添加粗体属性. BooleanDefaultTrue是设置b属性的Docx4j对象, 严格
	 * 来说我们不需要将值设置为true, 因为这是它的默认值.
	 */
	public static void addBoldStyle(RPr runProperties) {
		BooleanDefaultTrue b = new BooleanDefaultTrue();
		b.setVal(true);
		runProperties.setB(b);
	}

	/**
	 * 本方法像前面例子中一样再一次创建了普通的单元格
	 */
	public static void addNormalTableCell(Tr tableRow, String content) {
		Tc tableCell = factory.createTc();
		tableCell.getContent().add(
				wordMLPackage.getMainDocumentPart().createParagraphOfText(
						content));
		tableRow.getContent().add(tableCell);
	}

	/**
	 * 本方法给表格添加边框
	 */
	public static void addBorders(Tbl table, CTBorder border, TblBorders borders) {
		table.setTblPr(new TblPr());
		border.setColor("auto");
		border.setSz(new BigInteger("4"));
		border.setSpace(new BigInteger("0"));
		border.setVal(STBorder.SINGLE);

		borders.setBottom(border);
		borders.setLeft(border);
		borders.setRight(border);
		borders.setTop(border);
		borders.setInsideH(border);
		borders.setInsideV(border);
		table.getTblPr().setTblBorders(borders);
	}

	/**
	 * Docx4j拥有一个由字节数组创建图片部件的工具方法, 随后将其添加到给定的包中. 为了能将图片添加
	 * 到一个段落中, 我们需要将图片转换成内联对象. 这也有一个方法, 方法需要文件名提示, 替换文本,
	 * 两个id标识符和一个是嵌入还是链接到的指示作为参数.
	 * 一个id用于文档中绘图对象不可见的属性, 另一个id用于图片本身不可见的绘制属性. 最后我们将内联
	 * 对象添加到段落中并将段落添加到包的主文档部件.
	 *
	 * @param wordMLPackage 要添加图片的包
	 * @param bytes         图片对应的字节数组
	 * @throws Exception 不幸的createImageInline方法抛出一个异常(没有更多具体的异常类型)
	 */
	public static void addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes) throws Exception {
		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
		int docPrId = 1;
		int cNvPrId = 2;
		Inline inline = imagePart.createImageInline("Filename hint", "Alternative text", docPrId, cNvPrId, false);
		P paragraph = factory.createP();
		R run = factory.createR();
		paragraph.getContent().add(run);
		Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		wordMLPackage.getMainDocumentPart().addObject(paragraph);
	}

	/**
	 * 创建一个对象工厂并用它创建一个段落和一个可运行块R.
	 * 然后将可运行块添加到段落中. 接下来创建一个图画并将其添加到可运行块R中. 最后我们将内联
	 * 对象添加到图画中并返回段落对象.
	 *
	 * @param inline 包含图片的内联对象.
	 * @return 包含图片的段落
	 */
	public static P addInlineImageToParagraph(Inline inline, P paragraph, R run, Drawing drawing) {
		// 添加内联对象到一个段落中
		paragraph.getContent().add(run);
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		return paragraph;
	}

	/**
	 * 将图片从文件对象转换成字节数组.
	 *
	 * @param file 将要转换的文件
	 * @return 包含图片字节数据的字节数组
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static byte[] convertImageToByteArray(File file)
			throws FileNotFoundException, IOException {
		InputStream is = new FileInputStream(file);
		long length = file.length();
		// 不能使用long类型创建数组, 需要用int类型.
		if (length > Integer.MAX_VALUE) {
			System.out.println("File too large!!");
		}
		byte[] bytes = new byte[(int) length];
		int offset = 0;
		int numRead = 0;
		while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) {
			offset += numRead;
		}
		// 确认所有的字节都没读取
		if (offset < bytes.length) {
			System.out.println("Could not completely read file " + file.getName());
		}
		is.close();
		return bytes;
	}

	/**
	 * 创建table
	 *
	 * @param pt xlsx
	 * @return table
	 * @throws Exception 异常
	 */
	public static Tbl createTable(String pt) throws Exception {
		List<Map<String, Object>> mapList = ImportExcelUtil.parseExcel(pt);
		Tbl table = factory.createTbl();
		Map<String, Object> headData = mapList.get(0);
		Tr head = factory.createTr();
		for (Map.Entry<String, Object> entry : headData.entrySet()) {
			String key = entry.getKey();
			Tc tc = factory.createTc();
			tc.getContent().add(wordMLPackage.getMainDocumentPart().createParagraphOfText(key));
			head.getContent().add(tc);
		}
		table.getContent().add(head);
		mapList.forEach(d -> {
			Tr tr = factory.createTr();
			d.forEach((x, y) -> {
				Tc tc = factory.createTc();
				tc.getContent().add(wordMLPackage.getMainDocumentPart().createParagraphOfText(y.toString()));
				tr.getContent().add(tc);
			});
			table.getContent().add(tr);
		});
		return table;
	}

	/**
	 * 设置段落样式
	 *
	 * @param p       段落
	 * @param context 文本
	 * @return 段落
	 */
	public static P setStyle(P p, String context) {
		RPr rpr = factory.createRPr();
		RFonts font = new RFonts();
		//设置字体
		font.setAscii("宋体");
		font.setEastAsia("宋体");//经测试发现这个设置生效
		rpr.setRFonts(font);
		//设置颜色
		Color color = new Color();
		color.setVal("ABCDEF");
		rpr.setColor(color);
		//设置字体大小
		HpsMeasure fontSize = new HpsMeasure();
		fontSize.setVal(new BigInteger("48"));
		rpr.setSzCs(fontSize);
		rpr.setSz(fontSize);
		//设置粗体
		BooleanDefaultTrue bold = factory.createBooleanDefaultTrue();
		bold.setVal(Boolean.TRUE);
		rpr.setB(bold);
		//设置斜体
		BooleanDefaultTrue ltalic = new BooleanDefaultTrue();
		rpr.setI(ltalic);
		//设置删除线
		BooleanDefaultTrue deleteLine = new BooleanDefaultTrue();
		deleteLine.setVal(Boolean.TRUE);
		rpr.setStrike(deleteLine);
		//设置下划线
		U u = factory.createU();
		u.setVal(UnderlineEnumeration.SINGLE);
		u.setVal(UnderlineEnumeration.DOUBLE);//双下划线
		u.setVal(UnderlineEnumeration.DASH);//虚线
		u.setVal(UnderlineEnumeration.WAVE);//波浪线
		rpr.setU(u);
		//设置显示文本
		Text text = factory.createText();
		text.setValue(context);
		R r = factory.createR();
		r.getContent().add(text);
		r.setRPr(rpr);
		//设置段落居中
		PPr ppr = new PPr();
		Jc jc = new Jc();
		jc.setVal(JcEnumeration.LEFT);
		jc.setVal(JcEnumeration.RIGHT);
		jc.setVal(JcEnumeration.CENTER);
		ppr.setJc(jc);
		p.setPPr(ppr);
		p.getContent().add(r);
		return p;
	}

	/**
	 * 合并docx
	 *
	 * @param streams 输入流
	 * @return 合成流
	 * @throws Docx4JException 异常
	 * @throws IOException     异常
	 */
	public static InputStream mergeDocx(final List<InputStream> streams) throws Docx4JException, IOException {
		WordprocessingMLPackage target = null;
		final File generated = File.createTempFile("generated", ".docx");
		int chunkId = 0;
		Iterator<InputStream> it = streams.iterator();
		while (it.hasNext()) {
			InputStream is = it.next();
			if (is != null) {
				if (target == null) {
					// 复制第一个文档作为模板
					OutputStream os = new FileOutputStream(generated);
					os.write(IOUtils.toByteArray(is));
					os.close();
					target = WordprocessingMLPackage.load(generated);
				} else {
					// 添加其他文档
					insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(is), chunkId++);
				}
			}
		}
		if (target != null) {
			target.save(generated);
			return new FileInputStream(generated);
		} else {
			return null;
		}
	}

	// 插入文档
	private static void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {
		try {
			AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
					new PartName("/part" + chunkId + ".docx"));
			// afiPart.setContentType(new ContentType(CONTENT_TYPE));
			afiPart.setBinaryData(bytes);
			Relationship altChunkRel = main.addTargetPart(afiPart);
			CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
			chunk.setId(altChunkRel.getId());
			main.addObject(chunk);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

    /**
     * 根据输入流生成新的docx
	 * @param in 输入流
	 * @param path 生成目的地
	 * @return 生成路径
	 */
	public static String createDocx(InputStream in, String path) {
		if (in == null) {
			return null;
		}
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(path);
			byte[] bytes = new byte[1024];
			int index = 0;
			while ((index = in.read(bytes)) != -1) {
				out.write(bytes, 0, index);
				out.flush();
			}
			out.close();
			in.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}finally {
			try {
				if (out != null) {
					out.close();
				}
				if (in != null) {
					in.close();
				}
			}catch (IOException e){
				e.printStackTrace();
			}
		}
		return path;
	}

	/**
	 * 创建分页符
	 * @return
	 */
	public static P getPageBreak() {
		P p = new P();
		R r = new R();
		Br br = new Br();
		br.setType(STBrType.PAGE);
		r.getContent().add(br);
		p.getContent().add(r);
		return p;
	}
}
