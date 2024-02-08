package ser;


import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IWorkbasket;
import com.ser.blueline.metaDataComponents.*;
import com.ser.foldermanager.IFolder;
import com.ser.foldermanager.INode;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import jakarta.activation.DataHandler;
import jakarta.activation.DataSource;
import jakarta.activation.FileDataSource;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Utils {

    public static JSONObject getSystemConfig(ISession ses, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = ses.getDocumentServer().getStringMatrix("CCM_SYSTEM_CONFIG", ses);
        }
        if(mtrx == null) throw new Exception("SystemConfig Global Value List not found");

        List<List<String>> rawTable = mtrx.getRawRows();

        String srvn = ses.getSystem().getName().toUpperCase();
        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            String name = line.get(0);
            if(!name.toUpperCase().startsWith(srvn + ".")){continue;}
            name = name.substring(srvn.length() + ".".length());
            rtrn.put(name, line.get(1));
        }
        return rtrn;
    }
    static IRepresentation updateRepresentation(IDocument document, String type, String desc, String path) throws Exception {
        IRepresentation[] list = document.getRepresentationList();
        IRepresentation rtrn;
        for(IRepresentation repr : list){
            if(!repr.getType().equals(type)){continue;}
            if(!repr.getDescription().equals(desc)){continue;}

            document.removeRepresentation(repr.getRepresentationNumber());
        }

        rtrn = document.addRepresentation(type, desc);
        rtrn.addPartDocument(path);
        return rtrn;
    }
    static JSONObject getMainDocReviewStatuses(ISession ses, IDocumentServer srv, String prjn) throws Exception {
        IStringMatrix mtrx = getMainDocReviewStatusMatrix(ses, srv);
        if(mtrx == null) throw new Exception("MainDoc Review Status Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            if(!line.get(0).equals(prjn)){continue;}
            rtrn.put(line.get(1), line.get(2));
        }
        return rtrn;
    }
    static JSONObject getIssueStatuses(ISession ses, IDocumentServer srv, String prjn) throws Exception {
        IStringMatrix mtrx = getIssueStatusMatrix(ses, srv);
        if(mtrx == null) throw new Exception("Issue Status Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            if(!line.get(0).equals(prjn)){continue;}
            rtrn.put(line.get(1), line.get(2));
        }
        return rtrn;
    }
    static List<JSONObject> getWorkbaskets(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> rtrn = new ArrayList<>();
        IStringMatrix mtrx = getWorkbasketMatrix(ses, srv);
        String[] usrs = users.split("\\;");

        for (String usr : usrs) {
            JSONObject wusr = getWorkbasket(ses, srv, usr.trim(), mtrx);
            if(wusr == null){continue;}
            rtrn.add(wusr);
        }
        return rtrn;
    }
    static String getWorkbasketEMails(ISession ses, IDocumentServer srv, IBpmService bpm, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("ID") == null){continue;}
            IWorkbasket wb = bpm.getWorkbasket(wrba.getString("ID"));
            if(wb == null){continue;}
            String mail = wb.getNotifyEMail();
            if(mail == null){continue;}
            rtrn.add(mail);
        }
        return String.join(";", rtrn);
    }
    static String getWorkbasketDisplayNames(ISession ses, IDocumentServer srv, String users) throws Exception {
        List<JSONObject> wrbs = getWorkbaskets(ses, srv, users);
        List<String> rtrn = new ArrayList<>();
        for (JSONObject wrba : wrbs) {
            if(wrba.get("DisplayName") == null){continue;}
            rtrn.add(wrba.getString("DisplayName"));
        }
        return String.join(";", rtrn);
    }
    public static void sendHTMLMail(ISession ses, IDocumentServer srv, String mtpn, JSONObject pars) throws Exception {
        JSONObject mcfg = Utils.getMailConfig(ses, srv, mtpn);

        String host = mcfg.getString("host");
        String port = mcfg.getString("port");
        String protocol = mcfg.getString("protocol");
        String sender = mcfg.getString("sender");
        String subject = "";
        String mailTo = "";
        String mailCC = "";
        String attachments = "";

        if(pars.has("From")){
            sender = pars.getString("From");
        }
        if(pars.has("To")){
            mailTo = pars.getString("To");
        }
        if(pars.has("CC")){
            mailCC = pars.getString("CC");
        }
        if(pars.has("Subject")){
            subject = pars.getString("Subject");
        }
        if(pars.has("AttachmentPaths")){
            attachments = pars.getString("AttachmentPaths");
        }


        Properties props = new Properties();
        props.put("mail.debug","true");
        props.put("mail.smtp.debug", "true");

        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", port);

        String start_tls = (mcfg.has("start_tls") ? mcfg.getString("start_tls") : "");
        if(start_tls.equals("true")) {
            props.put("mail.smtp.starttls.enable", start_tls);
        }

        String auth = mcfg.getString("auth");
        props.put("mail.smtp.auth", auth);
        Authenticator authenticator = null;
        if(!auth.equals("false")) {
            String auth_username = mcfg.getString("auth.username");
            String auth_password = mcfg.getString("auth.password");

            if (host.contains("gmail")) {
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            if (protocol != null && protocol.contains("TLSv1.2"))  {
                props.put("mail.smtp.ssl.protocols", protocol);
                props.put("mail.smtp.ssl.trust", "*");
                props.put("mail.smtp.socketFactory.port", port);
                props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
                props.put("mail.smtp.socketFactory.fallback", "false");
            }
            authenticator = new Authenticator(){
                @Override
                protected PasswordAuthentication getPasswordAuthentication(){
                    return new PasswordAuthentication(auth_username, auth_password);
                }
            };
        }

        if(!Objects.equals(mailTo, "")) {
            Session session = (authenticator == null ? Session.getDefaultInstance(props) : Session.getDefaultInstance(props, authenticator));

            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(sender.replace(";", ",")));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(mailTo.replace(";", ",")));
            message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(mailCC.replace(";", ",")));
            message.setSubject(subject);

            Multipart multipart = new MimeMultipart("mixed");

            BodyPart htmlBodyPart = new MimeBodyPart();
            htmlBodyPart.setContent(getFileContent(pars.getString("BodyHTMLFile")), "text/html"); //5
            multipart.addBodyPart(htmlBodyPart);

            String[] atchs = attachments.split("\\;");
            for (String atch : atchs) {
                if (atch.isEmpty()) {
                    continue;
                }
                BodyPart attachmentBodyPart = new MimeBodyPart();
                attachmentBodyPart.setDataHandler(new DataHandler((DataSource) new FileDataSource(atch)));

                String fnam = Paths.get(atch).getFileName().toString();
                if (pars.has("AttachmentName." + fnam)) {
                    fnam = pars.getString("AttachmentName." + fnam);
                }

                attachmentBodyPart.setFileName(fnam);
                multipart.addBodyPart(attachmentBodyPart);

            }

            message.setContent(multipart);
            Transport.send(message);
        }else {

        }
    }
    static IStringMatrix getMailConfigMatrix(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        //IStringMatrix rtrn = srv.getStringMatrix("MailConfig" + (!mtpn.isEmpty() ? "." + mtpn : ""), ses);
        IStringMatrix rtrn = srv.getStringMatrix(mtpn, ses);
        if (rtrn == null) throw new Exception("MailConfig Global Value List not found:" + mtpn);
        return rtrn;
    }
    static String getFileContent (String path) throws Exception {
        return new String(Files.readAllBytes(Paths.get(path)));
    }
    static JSONObject getMailConfig(ISession ses, IDocumentServer srv, String mtpn) throws Exception {
        return getMailConfig(ses, srv, mtpn, null);
    }
    static JSONObject getMailConfig(ISession ses, IDocumentServer srv, String mtpn, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = getMailConfigMatrix(ses, srv, mtpn);
        }
        if(mtrx == null) throw new Exception("MailConfig Global Value List not found:" + mtpn);
        List<List<String>> rawTable = mtrx.getRawRows();

        JSONObject rtrn = new JSONObject();
        for(List<String> line : rawTable) {
            rtrn.put(line.get(0), line.get(1));
        }
        return rtrn;
    }
    static IStringMatrix getMainDocReviewStatusMatrix(ISession ses, IDocumentServer srv) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrix("MainDocReviewStatus", ses);
        if (rtrn == null) throw new Exception("MainDocReviewStatus Global Value List not found");
        return rtrn;
    }
    static IStringMatrix getIssueStatusMatrix(ISession ses, IDocumentServer srv) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrix("CCM_QCON_ISSUE-STATUSES", ses);
        if (rtrn == null) throw new Exception("IssueStatus Global Value List not found");
        return rtrn;
    }
    static IStringMatrix getWorkbasketMatrix(ISession ses, IDocumentServer srv) throws Exception {
        IStringMatrix rtrn = srv.getStringMatrixByID("Workbaskets", ses);
        if (rtrn == null) throw new Exception("Workbaskets Global Value List not found");
        return rtrn;
    }
    static JSONObject getWorkbasket(ISession ses, IDocumentServer srv, String userID) throws Exception {
        return getWorkbasket(ses, srv, userID, null);
    }
    public static IDocument getTemplateDocument(String prjNo, String tpltName, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.Template).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjNo).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.ObjectNumberExternal).append(" = '").append(tpltName).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = helper.createQuery(new String[]{Conf.Databases.Company} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    static JSONObject getWorkbasket(ISession ses, IDocumentServer srv, String userID, IStringMatrix mtrx) throws Exception {
        if(mtrx == null){
            mtrx = getWorkbasketMatrix(ses, srv);
        }
        if(mtrx == null) throw new Exception("Workbaskets Global Value List not found");
        List<List<String>> rawTable = mtrx.getRawRows();

        for(List<String> line : rawTable) {
            if(line.contains(userID)) {
                JSONObject rtrn = new JSONObject();
                rtrn.put("ID", line.get(0));
                rtrn.put("Name", line.get(1));
                rtrn.put("DisplayName", line.get(2));
                rtrn.put("Active", line.get(3));
                rtrn.put("Visible", line.get(4));
                rtrn.put("Type", line.get(5));
                rtrn.put("Organization", line.get(6));
                rtrn.put("Access", line.get(7));
                return rtrn;
            }
        }
        return null;
    }
    static void copyFile(String spth, String tpth) throws Exception {
        FileUtils.copyFile(new File(spth), new File(tpth));
    }
    public static void deleteRow(Sheet sheet, int rowNo) throws IOException {
        int lastRowNum = sheet.getLastRowNum();
        if (rowNo >= 0 && rowNo < lastRowNum) {
            sheet.shiftRows(rowNo + 1, lastRowNum, -1);
        }
        if (rowNo == lastRowNum) {
            Row removingRow=sheet.getRow(rowNo);
            if(removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
    public static void removeRows(String spth, String tpth, Integer shtIx, String prfx, Integer colIx, List<Integer> hlst, List<String> tlst) throws IOException {

        FileInputStream tist = new FileInputStream(spth);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheetAt(shtIx);
        JSONObject rows = Utils.getRowGroups(tsht, prfx, colIx);

        for (String pkey : rows.keySet()) {
            Row crow = (Row) rows.get(pkey);
            crow.getCell(colIx).setBlank();

            if(tlst.contains(pkey)){
                continue;
            }

            crow.setZeroHeight(true);
            //deleteRow(tsht, crow.getRowNum());
        }

        for(Integer hcix : hlst){
            tsht.setColumnHidden(hcix, true);
        }

        FileOutputStream tost = new FileOutputStream(tpth);
        twrb.write(tost);
        tost.close();

    }
    public static String saveDocReviewExcel(String templatePath, Integer shtIx, String tpltSavePath, JSONObject pbks) throws IOException {

        FileInputStream tist = new FileInputStream(templatePath);
        XSSFWorkbook twrb = new XSSFWorkbook(tist);

        Sheet tsht = twrb.getSheetAt(shtIx);
        for (Row trow : tsht){
            for(Cell tcll : trow){
                if(tcll.getCellType() != CellType.STRING){continue;}
                String clvl = tcll.getRichStringCellValue().getString();
                String clvv = updateCell(clvl, pbks);
                if(!clvv.equals(clvl)){
                    tcll.setCellValue(clvv);
                }

                if(clvv.indexOf("[[") != (-1) && clvv.indexOf("]]") != (-1)
                        && clvv.indexOf("[[") < clvv.indexOf("]]")){
                    String znam = clvv.substring(clvv.indexOf("[[") + "[[".length(), clvv.indexOf("]]"));
                    if(pbks.has(znam)){
                        tcll.setCellValue(znam);
                        String lurl = pbks.getString(znam);
                        if(!lurl.isEmpty()) {
                            Hyperlink link = twrb.getCreationHelper().createHyperlink(HyperlinkType.URL);
                            link.setAddress(lurl);
                            tcll.setHyperlink(link);
                        }
                    }
                }
            }
        }
        FileOutputStream tost = new FileOutputStream(tpltSavePath);
        twrb.write(tost);
        tost.close();
        return tpltSavePath;
    }

  public static String convertExcelToPdf(String excelPath, String pdfPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        workbook.getConverterSetting().setSheetFitToPage(true);
        workbook.saveToFile(pdfPath, FileFormat.PDF);

        return pdfPath;
    }

    public static String convertExcelToHtml(String excelPath, String htmlPath)  {
        Workbook workbook = new Workbook();
        workbook.loadFromFile(excelPath);
        Worksheet sheet = workbook.getWorksheets().get(0);
        HTMLOptions options = new HTMLOptions();
        options.setImageEmbedded(true);
        sheet.saveToHtml(htmlPath, options);
        return htmlPath;
    }
    public static String zipFiles(String zipPath, String tpltSavePath, List<String> expFilePaths) throws IOException {
        ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(new File(zipPath)));
        if(!tpltSavePath.isEmpty()) {
            //ZipEntry zltp = new ZipEntry("00." + Paths.get(tpltSavePath).getFileName().toString());
            ZipEntry zltp = new ZipEntry("_Transmittal." + FilenameUtils.getExtension(tpltSavePath));
            zout.putNextEntry(zltp);
            byte[] zdtp = Files.readAllBytes(Paths.get(tpltSavePath));
            zout.write(zdtp, 0, zdtp.length);
            zout.closeEntry();
        }

        for (String expFilePath : expFilePaths) {
            String fileName = Paths.get(expFilePath).getFileName().toString();
            fileName = fileName.replace("[@SLASH]", "/");
            ZipEntry zlin = new ZipEntry(fileName);

            zout.putNextEntry(zlin);
            byte[] zdln = Files.readAllBytes(Paths.get(expFilePath));
            zout.write(zdln, 0, zdln.length);
            zout.closeEntry();
        }
        zout.close();
        return zipPath;
    }
    public static String nameDocument(IDocument document) throws Exception {
        IDocumentPart partDocument = document.getPartDocument(document.getDefaultRepresentation() , 0);
        return partDocument.getFilename();
    }
    public static String exportDocument(IDocument document, String exportPath, String fileName) throws IOException {
        String rtrn ="";
        IDocumentPart partDocument = document.getPartDocument(document.getDefaultRepresentation() , 0);
        String fName = (!fileName.isEmpty() ? fileName : partDocument.getFilename());
        fName = fName.replaceAll("[\\\\/:*?\"<>|]", "_");
        try (InputStream inputStream = partDocument.getRawDataAsStream()) {
            IFDE fde = partDocument.getFDE();
            if (fde.getFDEType() == IFDE.FILE) {
                rtrn = exportPath + "/" + fName + "." + ((IFileFDE) fde).getShortFormatDescription();

                try (FileOutputStream fileOutputStream = new FileOutputStream(rtrn)){
                    byte[] bytes = new byte[2048];
                    int length;
                    while ((length = inputStream.read(bytes)) > -1) {
                        fileOutputStream.write(bytes, 0, length);
                    }
                }
            }
        }
        return rtrn;
    }
    public static INode getNode(IFolder fold, String fldn){
        List<INode> nodesByName = fold.getNodesByName(fldn);
        return fold.getNodeByID(nodesByName.get(0).getID());
    }
    public static String getCellValue(Sheet sheet, String refn){

        CellReference cr = new CellReference(refn);
        Row row = sheet.getRow(cr.getRow());
        Cell rtrn = row.getCell(cr.getCol());
        return rtrn.getRichStringCellValue().getString();
    }

    public static String updateCell(String str, JSONObject bookmarks){
        StringBuffer rtr1 = new StringBuffer();
        String tmp = str + "";
        Pattern ptr1 = Pattern.compile( "\\{([\\w\\.]+)\\}" );
        Matcher mtc1 = ptr1.matcher(tmp);
        while(mtc1.find()) {
            String mk = mtc1.group(1);
            String mv = "";
            if(bookmarks.has(mk)){
                mv = bookmarks.getString(mk);
            }
            mtc1.appendReplacement(rtr1,  mv);
        }
        mtc1.appendTail(rtr1);
        tmp = rtr1.toString();

        return tmp;
    }
    static IInformationObject[] getSubProcessies(String mainDocId, ProcessHelper helper)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.SubProcess).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.MainTaskReference).append(" = '").append(mainDocId).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        return helper.createQuery(new String[]{Conf.Databases.BPM} , whereClause, 0);
    }
    public static void saveFileContent(String path, String cntn) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(path);
        byte[] bctn = cntn.getBytes();
        outputStream.write(bctn);

        outputStream.close();
    }
    public static JSONObject getRowGroups(Sheet sheet, String prfx, Integer colIx)  {
        JSONObject rtrn = new JSONObject();
        for (Row row : sheet) {
            Cell cll1 = row.getCell(colIx);
            if(cll1 == null){continue;}

            String cval = cll1.getRichStringCellValue().getString();
            if(cval.isEmpty()){continue;}

            if(!cval.startsWith("[&" + prfx + ".")
                    || !cval.endsWith("&]")){continue;}

            String znam = cval.substring(("[&" + prfx + ".").length(), cval.length() - ("]&").length());
            rtrn.put(znam, row);

        }
        return rtrn;
    }
    public static IQueryDlg findQueryDlgForQueryClass(IQueryClass queryClass) {
        IQueryDlg dlg = null;
        if (queryClass != null) {
            dlg = queryClass.getQueryDlg("default");
        }

        return dlg;
    }
    public static IQueryParameter query(ISession session, IQueryDlg queryDlg, Map<String, String> descriptorValues) {
        IDocumentServer documentServer = session.getDocumentServer();
        ISerClassFactory classFactory = documentServer.getClassFactory();
        IQueryParameter queryParameter = null;
        IQueryExpression expression = null;
        IComponent[] components = queryDlg.getComponents();

        for(int i = 0; i < components.length; ++i) {
            if (components[i].getType() == IMaskedEdit.TYPE) {
                IControl control = (IControl)components[i];
                String descriptorId = control.getDescriptorID();
                String value = (String)descriptorValues.get(descriptorId);
                if (value != null && value.trim().length() > 0) {
                    IDescriptor descriptor = documentServer.getDescriptor(descriptorId, session);
                    IQueryValueDescriptor queryValueDescriptor = classFactory.getQueryValueDescriptorInstance(descriptor);
                    queryValueDescriptor.addValue(value);
                    IQueryExpression expr = queryValueDescriptor.getExpression();
                    if (expression != null) {
                        expression = classFactory.getExpressionInstance(expression, expr, 0);
                    } else {
                        expression = expr;
                    }
                }
            }
        }

        if (expression != null) {
            queryParameter = classFactory.getQueryParameterInstance(session, queryDlg, expression);
        }

        return queryParameter;
    }
    public static IDocumentHitList executeQuery(ISession session, IQueryParameter queryParameter) {
        IDocumentServer documentServer = session.getDocumentServer();
        return documentServer.query(queryParameter, session);
    }
    public static IDocument[] getEngCopyDocuments(ISession session, String docKey) throws IOException {
        String searchClassName = "Search Review Copy Documents";
        IDocumentServer documentServer = session.getDocumentServer();
        IDescriptor descriptor2 = documentServer.getDescriptorForName(session, Conf.Descriptors.MainDocumentID);
        IQueryClass queryClass = documentServer.getQueryClassByName(session, searchClassName);
        IQueryDlg queryDlg = findQueryDlgForQueryClass(queryClass);
        Map<String, String> searchDescriptors = new HashMap();
        searchDescriptors.put(descriptor2.getId(), docKey);
        IQueryParameter queryParameter = query(session, queryDlg, searchDescriptors);
        if (queryParameter != null) {
            IDocumentHitList hitresult = executeQuery(session, queryParameter);
            IDocument[] hits = hitresult.getDocumentObjects();
            queryParameter.close();
            return hits != null && hits.length > 0 ? hits : null;
        } else {
            return null;
        }
    }
}
