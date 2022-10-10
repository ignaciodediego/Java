package com.urbanitae.shareholders.service.impl;

import com.google.common.collect.Lists;
import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import com.urbanitae.commons.advice.UrbanitaeBusinessException;
import com.urbanitae.commons.constants.UrbanitaeExceptionMessages;
import com.urbanitae.shareholders.model.*;
import com.urbanitae.shareholders.model.organization.Organization;
import com.urbanitae.shareholders.model.project.Project;
import com.urbanitae.shareholders.model.project.UserParticipation;
import com.urbanitae.shareholders.model.user.PersonalData;
import com.urbanitae.shareholders.model.user.User;
import com.urbanitae.shareholders.model.wallet.Wallet;
import com.urbanitae.shareholders.repositories.MeetingUserAnswerRepository;
import com.urbanitae.shareholders.service.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.joda.time.LocalDate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.util.StringUtils;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.HttpServerErrorException;
import org.springframework.web.client.RestTemplate;

import javax.management.InstanceNotFoundException;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.Function;
import java.util.stream.Collectors;

@Component
@Slf4j
public class MeetingUserAnswerServiceDefault implements MeetingUserAnswerService {

    @Autowired
    private MeetingUserAnswerRepository meetingUserAnswerRepository;

    @Autowired
    private MeetingUserAnswerService meetingUserAnswerService;

    @Autowired
    private ShareHolderMeetingService shareHolderMeetingService;

    @Autowired
    private ProjectUploadFile projectUploadFile;

    @Autowired
    private ProjectService projectService;

    @Autowired
    private WalletService walletService;

    @Autowired
    private UserService userService;

    @Autowired
    private OrganizationService organizationService;

    @Value("#{${meeting.answers}}")
    private Map<String, String> answerKey;

    @Autowired
    RestTemplate restTemplate;

    ExecutorService executor = Executors.newSingleThreadExecutor();

    @Override
    public MeetingUserAnswer saveMeetingUserAnswer(MeetingUserAnswer meetingUserAnswer)
            throws UrbanitaeBusinessException, FileNotFoundException, DocumentException, InstanceNotFoundException {
        ShareHolderMeeting shareHolderMeeting = shareHolderMeetingService
                .getShareHolderMeeting(meetingUserAnswer.getMeetingId());
        if (shareHolderMeeting.getStatus() == MeetingStatus.ACTIVE) {

            Investor investor = shareHolderMeeting.getInvited().stream().filter(
                    investorElement -> investorElement.getUserId().compareTo(meetingUserAnswer.getUserId()) == 0)
                    .findFirst().orElse(null);
            if (investor != null) {
                if (investor.getCompleteMeeting()) {
                    throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.USER_YET_ANSWER_ERROR,
                            Lists.newArrayList(""), UrbanitaeExceptionMessages.USER_YET_ANSWER_ERROR_CODE);
                } else {
                    MeetingUserAnswer meetingSave = null;
                    switch (meetingUserAnswer.getVoteType()) {
                        case "online":
                            if (meetingUserAnswer.getQuestionAnswerList() == null || meetingUserAnswer
                                    .getQuestionAnswerList().size() != shareHolderMeeting.getQuestions().size()) {
                                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.MEETING_NOT_COMPLETE_ERROR,
                                        Lists.newArrayList(""), UrbanitaeExceptionMessages.MEETING_NOT_COMPLETE_ERROR_CODE);
                            }
                            break;
                        case "delegate_admin":
                            if (meetingUserAnswer.getDelegate() != null
                                    || (meetingUserAnswer.getQuestionAnswerList() != null
                                    && !meetingUserAnswer.getQuestionAnswerList().isEmpty())) {
                                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR,
                                        Lists.newArrayList(""), UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR_CODE);
                            }
                            break;
                        case "delegate_third":
                            if (meetingUserAnswer.getDelegate() == null
                                    || (meetingUserAnswer.getQuestionAnswerList() != null
                                    && !meetingUserAnswer.getQuestionAnswerList().isEmpty())) {
                                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR,
                                        Lists.newArrayList(""), UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR_CODE);
                            }
                            break;

                    }
                    investor.setCompleteMeeting(Boolean.TRUE);
                    meetingUserAnswer.setId(meetingUserAnswer.getUserId() + "_" + meetingUserAnswer.getMeetingId());
                    try {
                        meetingSave = meetingUserAnswerRepository.insert(meetingUserAnswer);

                    } catch (org.springframework.dao.DuplicateKeyException e) {
                        log.error("duplicate request");
                        return meetingUserAnswer;
                    }
                    checkMeetingClose(meetingUserAnswer, shareHolderMeeting);
                    log.info("check send mail {} -- {}", meetingUserAnswer.getVoteType(), meetingUserAnswer);
                    if (meetingUserAnswer.getVoteType().compareTo("delegate_third") == 0) {
                        sendDelegateVoteMail(meetingUserAnswer);
                    }

                    checkMeetingClose(meetingUserAnswer, shareHolderMeeting);
                    shareHolderMeetingService.save(shareHolderMeeting);
                    return meetingSave;

                }
            } else {
                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR,
                        Lists.newArrayList(""), UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR_CODE);
            }
        } else {
            throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.MEETING_NOT_ACTIVE_ERROR,
                    Lists.newArrayList(""), UrbanitaeExceptionMessages.MEETING_NOT_ACTIVE_ERROR_CODE);
        }
    }

    private void sendDelegateVoteMail(MeetingUserAnswer meetingUserAnswer)
            throws FileNotFoundException, DocumentException, InstanceNotFoundException, UrbanitaeBusinessException {

        ShareHolderMeeting meeting = this.shareHolderMeetingService
                .getShareHolderMeeting(meetingUserAnswer.getMeetingId());

        try {

            User user = this.userService.findById(meetingUserAnswer.getUserId());
            Project project = projectService.getProjectById(meeting.getProjectId());
            Wallet wallet = this.walletService.getProjectWallet(project.getWallet());

            Map<String, Object> mailNotification = new HashMap<>();

            mailNotification.put("mailTo", meetingUserAnswer.getDelegate().getEmail());
            mailNotification.put("subject", "meeting.delegate.document");
            if (StringUtils.isEmpty(user.getOrganization())) {
                mailNotification.put("template", "meeting.delegate.document");
            } else {
                mailNotification.put("template", "saas_meeting.delegate.document");
            }

            Map<String, Object> args = new HashMap<>();
            args.put("body", "Documentación de usuarios para el notario");
            args.put("name", user.getPersonalData().getFirstName());
            args.put("delegateName", meetingUserAnswer.getDelegate().getName());
            args.put("legalName", wallet.getCompanyInfo().getSocialReason());

            SimpleDateFormat dt1 = new SimpleDateFormat("dd/MM/yyyyy, hh:mm  ");

            args.put("meetingDate", dt1.format(meeting.getMeetingDate()));
            args.put("meetingAddress", meeting.getMeetingAddress());
            args.put("meetingName", meeting.getName());
            if (StringUtils.isEmpty(user.getOrganization())) {
                Organization organization = this.organizationService.getOrganization(user.getOrganization());
                args.put("contactPhone", organization.getContactPhone());
                args.put("contactEmail", organization.getContactEmail());
                args.put("organization", organization.getKey());
            }
            mailNotification.put("mailMap", args);

            MultiValueMap<String, Object> multipartRequest = new LinkedMultiValueMap<>();

            HttpHeaders jsonHeader = new HttpHeaders();
            jsonHeader.setContentType(MediaType.APPLICATION_JSON);
            HttpEntity<Map<String, Object>> mailNotificationPart = new HttpEntity<>(mailNotification, jsonHeader);
            multipartRequest.add("mail", mailNotificationPart);

            ByteArrayResource autorizationFile = generateAutorizationFile(meetingUserAnswer, user, wallet, meeting);

            HttpHeaders fileHeeaders = new HttpHeaders();
            fileHeeaders.setContentDispositionFormData("attachment", "autorizacion.pdf");
            HttpEntity<ByteArrayResource> filePart = new HttpEntity<>(autorizationFile, fileHeeaders);

            multipartRequest.add("attachment", filePart);

            try {
                log.info("got to send mail {}", args);
                restTemplate.postForLocation("http://mail/mail", multipartRequest);
                log.info("mail yet send ");
            } catch (HttpClientErrorException | HttpServerErrorException exception) {
                log.error("can not send mail {}", ExceptionUtils.getStackTrace(exception));
            }

        } catch (Exception e) {
            log.error("can not generate and send mail {}", ExceptionUtils.getStackTrace(e));
        }
    }

    private ByteArrayResource generateAutorizationFile(MeetingUserAnswer meetingUserAnswer, User user, Wallet wallet,
                                                       ShareHolderMeeting meeting) throws DocumentException, FileNotFoundException {

        String fileName = "autorization.pdf";

        Document document = new Document();
        ByteArrayOutputStream byteOutput = new ByteArrayOutputStream();
        PdfWriter.getInstance(document, byteOutput);

        document.open();
        Font font = FontFactory.getFont(FontFactory.HELVETICA, 14, BaseColor.BLACK);
        Paragraph chunk = new Paragraph("REPRESENTACIÓN EN JUNTA DE SOCIOS.\n\n", font);
        Paragraph firstParagraph = new Paragraph("Por medio de la presente D." + user.getPersonalData().getFirstName()
                + " " + user.getPersonalData().getLastName() + ", con número de documento "
                + user.getPersonalData().getIdNumber() + ", en su calidad de socio/a de "
                + wallet.getCompanyInfo().getSocialReason() + " SL.\n\n", font);

        SimpleDateFormat dt1 = new SimpleDateFormat("dd/MM/yyyyy, hh:mm");
        String startFormatDate = dt1.format(meeting.getMeetingDate());

        Paragraph secondParagraph = new Paragraph("AUTORIZA Y CONFIERE a D. "
                + meetingUserAnswer.getDelegate().getName() + " " + meetingUserAnswer.getDelegate().getSurname()
                + " con número de documento " + meetingUserAnswer.getDelegate().getIdNumber()
                + " para que actuando en su nombre, le represente de la forma más plena y eficaz posible en derecho, en la "
                + meeting.getName() + " convocada para el día " + startFormatDate + ", en la dirección "
                + meeting.getMeetingAddress() + ", con el fin de tratar los siguientes puntos del Orden el día:\n\n",
                font);

        document.add(chunk);
        document.add(firstParagraph);
        document.add(secondParagraph);

        for (int i = 0; i < meeting.getQuestions().size(); i++) {
            Paragraph question = new Paragraph(
                    "     " + (i + 1) + ". " + meeting.getQuestions().get(i).getQuestion() + ".\n\n", font);
            document.add(question);
        }

        Paragraph endParagraph = new Paragraph("Se concede al citado representante la facultad para asistir a la "
                + meeting.getName()
                + ", aceptar su constitución y la inclusión de otros puntos en su orden del día, así como para deliberar y proponer.\n",
                font);
        document.add(endParagraph);
        LocalDate today = LocalDate.now();

        Paragraph dateParagraph = new Paragraph("En Madrid a " + today.getYear() + "\n\n\n", font);
        document.add(dateParagraph);

        Paragraph firmParagraph1 = new Paragraph("................\n", font);
        document.add(firmParagraph1);
        Paragraph firmParagraph2 = new Paragraph("Firmado D." + user.getPersonalData().getFirstName() + " "
                + user.getPersonalData().getLastName() + "\n", font);
        document.add(firmParagraph2);
        Paragraph firmParagraph3 = new Paragraph("Socio de " + wallet.getCompanyInfo().getSocialReason() + "\n\n\n",
                font);
        document.add(firmParagraph3);

        firmParagraph1 = new Paragraph("................\n", font);
        document.add(firmParagraph1);
        firmParagraph2 = new Paragraph("Firmado D." + meetingUserAnswer.getDelegate().getName() + " "
                + meetingUserAnswer.getDelegate().getSurname() + "\n", font);
        document.add(firmParagraph2);
        firmParagraph3 = new Paragraph("Representante voluntaria \n", font);
        document.add(firmParagraph3);

        document.close();
        try {
            byteOutput.close();
        } catch (IOException e) {
            log.info("error writing file {} ", ExceptionUtils.getStackTrace(e));
        }

        return new ByteArrayResource(byteOutput.toByteArray(), fileName);

    }

    @Override
    public MeetingSummary getSummaryResult(String meetingId)
            throws InstanceNotFoundException, UrbanitaeBusinessException {
        MeetingSummary meetingSummary = meetingUserAnswerRepository.meetingSummaryResult(meetingId);

        ShareHolderMeeting meeting = this.shareHolderMeetingService.getShareHolderMeeting(meetingId);

        Project project = this.projectService.getProjectById(meeting.getProjectId());

        // filterComments
        log.info("view list {}", meeting.getInvited());
        Map<String, User> userMap = meeting.getInvited().stream()
                .filter(userInvited -> userInvited.getCompleteMeeting())
                .map(userInvited -> this.userService.findById(userInvited.getUserId()))
                .collect(Collectors.toMap(User::getId, Function.identity(), (first, second) -> first));

        meetingSummary.getQuestion().stream().forEach(question -> question.getResult().stream().forEach(result -> {

            checkComments(result);

            BigDecimal percentage = BigDecimal.ZERO;
            List<UserInvestor> userParticipationList = new ArrayList<>();
            for (String userId : result.getUserVoted()) {
                log.info("user answer {}", userId);

                User user = userMap.get(userId);

                Optional<UserParticipation> projectParticipation = project.getFundedInfo().getParticipations().stream()
                        .filter(userParticipation -> {
                            return userParticipation.getUserId().compareTo(user.getWalletId()) == 0;
                        })
                        .findFirst();
                percentage = percentage
                        .add(checkParticipation(projectParticipation));

                result.getComments().stream().filter(comment -> comment.getUserId().compareTo(user.getId()) == 0)
                        .forEach(comentuser -> comentuser.setUser(
                                user.getPersonalData().getFirstName() + " " + user.getPersonalData().getLastName()));

                userParticipationList.add(new UserInvestor(
                        user.getPersonalData().getFirstName() + " " + user.getPersonalData().getLastName(),
                        user.getPersonalData().getIdNumber(),
                        projectParticipation.isPresent() ? projectParticipation.get().getFunds() : new BigDecimal("0"),
                        projectParticipation.isPresent() ? projectParticipation.get().getCompanyPercentage()
                                : new BigDecimal("0")));
            }

            result.setUserList(userParticipationList);
            result.setPercentage(percentage);

        }));
        return meetingSummary;
    }

    private BigDecimal checkParticipation(Optional<UserParticipation> participation) {
        if (participation.isPresent() && participation.get().getCompanyPercentage() != null) {
            return participation.get().getCompanyPercentage();
        }
        return BigDecimal.ZERO;
    }

    private void checkComments(Result result) {
        if (result.getComments() != null) {
            List<QuestionComments> newComments = new ArrayList<>();
            for (QuestionComments comments : result.getComments()) {
                if (comments.getComment() != null && comments.getComment().length() != 0) {
                    newComments.add(comments);
                }
            }
            result.setComments(newComments);
        }
    }

    @Override
    public MeetingUserAnswer getMeetingUserAnswer(String meetingId, String userId) {
        return meetingUserAnswerRepository.findByUserIdIsAndMeetingIdIs(userId, meetingId);
    }

    @Override
    public List<MeetingUserAnswer> getMeetingAnswer(String meetingId) {
        return meetingUserAnswerRepository.findByMeetingIdIs(meetingId);
    }

    @Override
    public Resource generateExcelMeeting(String meetingId)
            throws InstanceNotFoundException {

        log.info("generateExcelMeeting meetingID {}", meetingId);

        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            ShareHolderMeeting meeting = this.shareHolderMeetingService.getShareHolderMeeting(meetingId);
            crearHojaUno(wb, meetingId);
            crearHojaDos(wb, meeting);

        } catch (Exception e) {

            log.error("No meetingId {}", ExceptionUtils.getStackTrace(e));
        }

        InputStream is = null;

        try {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            wb.close();
            byte[] barray = bos.toByteArray();
            is = new ByteArrayInputStream(barray);
            return new InputStreamResource(is);

        } catch (IOException e) {
            try {
                wb.close();
            } catch (Exception ex) {

                log.error("Error closing workbook {}", ExceptionUtils.getStackTrace(ex));
            }
            log.error("No ByteArrayOutputStream {}", ExceptionUtils.getStackTrace(e));
            return null;
        }

    }

    private void imprimirFila(XSSFWorkbook wb, XSSFSheet hoja, int numFila, List<String> cabecera, int posColumna, XSSFCellStyle estilo, Boolean setValue) {

        Row row = hoja.getRow(numFila) != null ? hoja.getRow(numFila) : hoja.createRow(numFila);

        try {
            for (int a = 0; a < cabecera.size(); a++) {

                Cell cell = row.createCell(posColumna);
                cell.setCellStyle(estilo);
                if (setValue) {
                    cell.setCellValue(cabecera.get(a));
                } else {
                    cell.setCellFormula(cabecera.get(a));
                }
                posColumna++;
            }
        } catch (Exception e) {
            System.out.println("Error creating row");
        }
    }

    private void imprimirFilaDouble(XSSFWorkbook wb, XSSFSheet hoja, int numFila, List<Double> cabecera, int posColumna, XSSFCellStyle estilo) {

        Row row = hoja.getRow(numFila) != null ? hoja.getRow(numFila) : hoja.createRow(numFila);

        try {
            for (int a = 0; a < cabecera.size(); a++) {

                Cell cell = row.createCell(posColumna);
                cell.setCellStyle(estilo);
                cell.setCellValue(cabecera.get(a));
                posColumna++;
            }
        } catch (Exception e) {
            System.out.println("Error creating row");
        }
    }

    private NumberFormat dameNumberFormat() {

        NumberFormat nf = NumberFormat.getNumberInstance(new Locale("es"));
        DecimalFormat df = (DecimalFormat) nf;
        df.setMaximumFractionDigits(2);
        return df;

    }

    private void crearHojaDos(XSSFWorkbook wb, ShareHolderMeeting meeting) {

        log.info("crearHojaDos meetingID {}", meeting.getId());

        ///Esto sacarlo en la optimizaciónde código
        XSSFSheet personas = wb.createSheet("Personas");
        Row rowPersonas = personas.createRow(0);
        Cell cellPersonas = rowPersonas.createCell(0);
        cellPersonas.setCellStyle(dameEstiloGeneral(wb, true, ""));
        cellPersonas.setCellValue("Gente que falta por votar");

        List<Investor> lInvestorNoComplete = meeting.getInvited().stream()
                .filter(invitado -> invitado.getCompleteMeeting().equals(false))
                .collect(Collectors.toList());

        int posNoComplete = 0;
        int posLinea = 1;

        for (int dd = 0; dd < lInvestorNoComplete.size(); dd++) {

            posNoComplete = 0;
            Row rowPersonasNoComplete = personas.createRow(posLinea);
            String idNoVotado = lInvestorNoComplete.get(dd).getUserId();

            try {

                User usuario = this.userService.findById(idNoVotado);
                if (usuario != null) {
                    Cell cell = rowPersonasNoComplete.createCell(posNoComplete);
                    cell.setCellStyle(dameEstiloGeneral(wb, false, ""));

                    cell.setCellValue(usuario.getPersonalData().getFirstName() + " " + usuario.getPersonalData().getLastName());
                    posNoComplete++;

                    Cell cell2 = rowPersonasNoComplete.createCell(posNoComplete);
                    cell2.setCellStyle(dameEstiloGeneral(wb, false, ""));
                    cell2.setCellValue(idNoVotado);
                    posNoComplete++;

                    Cell cell3 = rowPersonasNoComplete.createCell(posNoComplete);
                    cell3.setCellStyle(dameEstiloGeneral(wb, false, ""));
                    cell3.setCellValue(usuario.getMail() != null ? usuario.getMail() : "");

                    posLinea++;
                }

            } catch (Exception e) {
                log.error("No se puede encontrar el usuario {} Exception: {}", idNoVotado, ExceptionUtils.getStackTrace(e));
                //throw new InstanceNotFoundException("User " + idNoVotado + " does not exists");
            }
        }
        for (int i = 0; i < 13; i++) { //13 es el máx de columnas
            personas.autoSizeColumn(i);
        }

    }

    private void crearHojaUno(XSSFWorkbook wb, String meetingId) {


        log.info("crearHojaUno meetingID {}", meetingId);

        XSSFSheet hoja = wb.createSheet("Junta");

        try {
            ShareHolderMeeting meeting = this.shareHolderMeetingService.getShareHolderMeeting(meetingId);

            PaymentElements pElements = projectService.getPayment(meeting.getProjectId());
            List<UserParticipation> listUP = null;

            if (pElements != null) {
                listUP = pElements.getParticipations();
            }

            BigDecimal companyCapital = new BigDecimal(0);

            try {

                companyCapital = projectService.getProjectById(meeting.getProjectId()).getFund().getCompanyCapital();

            } catch (Exception e) {

                log.error("No company capital {}", ExceptionUtils.getStackTrace(e));
            }


            List<String> cabecera1 = Arrays.asList("Proyecto", meeting.getName() != null ? meeting.getName() : "", companyCapital != null ? dameNumberFormat().format(companyCapital) + "€" : "0", "Esto es la cantidad total del proyecto");

            imprimirFila(wb, hoja, 1, cabecera1, 1, dameEstiloGeneral(wb, true, ""), true);

            List<String> cabecera21 = Arrays.asList("VOTACIÓN WEB");
            List<String> cabecera22 = Arrays.asList("DELEGACIÓN VOTO WEB", "DELEGACIÓN VOTO CORREO", "PROMOTOR", "APROBACIÓN TOTAL");

            imprimirFila(wb, hoja, 3, cabecera21, 2, dameEstiloColor(wb, "verde", ""), true);
            hoja.addMergedRegion(new CellRangeAddress(3, 3, 2, 4));
            imprimirFila(wb, hoja, 3, cabecera22, 5, dameEstiloColor(wb, "verde", ""), true);

            List<String> cabecera3 = Arrays.asList("Si", "No", "Abstención", "Si", "Si", "Si", "");
            imprimirFila(wb, hoja, 4, cabecera3, 2, dameEstiloColor(wb, "gris", ""), true);

            int fila = 5;
            int filaInicioOcultos = 15;

            List<MeetingUserAnswer> listUserAnswer = meetingUserAnswerRepository.findByMeetingIdIs(meetingId);

            if (listUserAnswer.size() > 0) {

                List<Question> listaPreguntas = meeting.getQuestions();

                for (int f = 0; f < listaPreguntas.size(); f++) {

                    List<String> cabecera41 = Arrays.asList(listaPreguntas.get(f).getQuestion());
                    List<String> cabecera42 = Arrays.asList("O" + filaInicioOcultos, "O" + (filaInicioOcultos + 1), "O" + (filaInicioOcultos + 2), "O13", "O14", "0", "SUM(C" + (fila + 1) + ",F" + (fila + 1) + ",G" + (fila + 1) + ",H" + (fila + 1) + ")");
                    imprimirFila(wb, hoja, fila, cabecera41, 1, dameEstiloColor(wb, "gris", ""), true);
                    imprimirFila(wb, hoja, fila, cabecera42, 2, dameEstiloGeneral(wb, false, "percentage"), false);
                    fila++;
                    filaInicioOcultos = filaInicioOcultos + 3;
                }

                String optionEncabezado = "";
                String respuestaVoto = "";
                List<MeetingUserAnswer> listaUsuariosVotadoOnline = listUserAnswer.stream()
                        .filter(userAnswer -> userAnswer.getVoteType().equals("online"))
                        .collect(Collectors.toList());

                int hiddenPosition = 14;
                //BUCLE PARA CADA UNA DE LAS PREGUNTAS
                for (int g = 0; g < listaPreguntas.size(); g++) {
                    fila++;
                    for (int h = 0; h < 3; h++) {
                        //ENCABEZADO
                        if (h == 0) {
                            optionEncabezado = "VOTOS A FAVOR ";
                            respuestaVoto = "answer.yes";
                        } else if (h == 1) {
                            optionEncabezado = "VOTOS EN CONTRA ";
                            respuestaVoto = "answer.no";
                        } else if (h == 2) {
                            optionEncabezado = "ABSTENCIÓN ";
                            respuestaVoto = "answer.abstention";
                        }
                        List<String> encabezadoResp = Arrays.asList(optionEncabezado + listaPreguntas.get(g).getQuestion(), "Cantidad", "%");
                        List<MeetingUserAnswer> listVotoUserFiltrada = dameValorRespuesta(g, respuestaVoto, listaUsuariosVotadoOnline);

                        fila = imprimirTablaVotos(fila, wb, hoja, encabezadoResp, listVotoUserFiltrada, listUP, hiddenPosition, 1, companyCapital);
                        hiddenPosition++;
                        fila = fila + 2;
                    }
                    fila++;
                    Row filaNegra = hoja.createRow(fila);
                    Cell cellNegra = filaNegra.createCell(1);
                    cellNegra.setCellStyle(dameEstiloColor(wb, "negro", ""));
                    hoja.addMergedRegion(new CellRangeAddress(fila, fila, 1, 3));
                }
                //ACABA EL BUCLE DE LAS PREGUNTAS
            } else {

                List<String> cabeceraError = Arrays.asList("ERROR: el usuario no tiene preguntas respondidas en la tabla meetingUserAnswer");
                imprimirFila(wb, hoja, fila, cabeceraError, 1, dameEstiloColor(wb, "gris", ""), true);
            }

            fila = 14;

            List<String> cabeceraAdmin = Arrays.asList("DELEGACIÓN VOTO WEB EN ADMINISTRADOR", "Cantidad", "%");

            List<MeetingUserAnswer> listaUsuariosDelegadoAdmin = listUserAnswer.stream()
                    .filter(userAnswer -> userAnswer.getVoteType().equals("delegate_admin"))
                    .collect(Collectors.toList());

            fila = imprimirTablaVotos(fila, wb, hoja, cabeceraAdmin, listaUsuariosDelegadoAdmin, listUP, 12, 5, companyCapital);

            fila = fila + 2;

            List<String> cabeceraDelCEO = Arrays.asList("DELEGACIÓN VOTO CORREO EN EL CEO", "Cantidad", "%");

            List<MeetingUserAnswer> listaUsuariosDelegadoCeo = listUserAnswer.stream()
                    .filter(userAnswer -> userAnswer.getVoteType().equals("delegate_ceo"))
                    .collect(Collectors.toList());

            fila = imprimirTablaVotos(fila, wb, hoja, cabeceraDelCEO, listaUsuariosDelegadoCeo, listUP, 13, 5, companyCapital);

            fila = fila + 2;

            List<String> cabeceraAportacionPromotor = Arrays.asList("Aportación Promotor", "No lo sabemos", "0.00 %");
            imprimirFila(wb, hoja, fila, cabeceraAportacionPromotor, 5, dameEstiloColor(wb, "gris", ""), true);

            fila = fila + 3;

            List<String> cabeceraResultFinal1 = Arrays.asList("RESULTADO TOTAL DE LA JUNTA");
            List<String> cabeceraResultFinal2 = Arrays.asList("SUM(N13:N17)");
            List<String> cabeceraResultFinal3 = Arrays.asList("SUM(O13:O17)");

            imprimirFila(wb, hoja, fila, cabeceraResultFinal1, 5, dameEstiloColor(wb, "verde", ""), true);
            imprimirFila(wb, hoja, fila, cabeceraResultFinal2, 6, dameEstiloColor(wb, "verde", "moneda"), false);
            imprimirFila(wb, hoja, fila, cabeceraResultFinal3, 7, dameEstiloColor(wb, "verde", "percetange"), false);

            for (int i = 0; i < 13; i++) { //13 es el máx de columnas
                hoja.autoSizeColumn(i);
            }

            hoja.setColumnHidden(14, true);
            hoja.setColumnHidden(13, true);

        } catch (Exception e) {

            log.error("Error creating first Tag");
            List<String> cabeceraError = Arrays.asList("ERROR: al crear el excel, el Meeting no contiene toda la información necesaria.");
            imprimirFila(wb, hoja, 1, cabeceraError, 1, dameEstiloColor(wb, "gris", ""), true);
            for (int i = 0; i < 13; i++) { //13 es el máx de columnas
                hoja.autoSizeColumn(i);
            }

        }
    }

    private int imprimirTablaVotos(int fila, XSSFWorkbook wb, XSSFSheet hoja, List<String> cabecera, List<MeetingUserAnswer> listaUsuarios, List<UserParticipation> listUP, int posFilaOculta, int posColumna, BigDecimal companyCapital) {

        imprimirFila(wb, hoja, fila, cabecera, posColumna, dameEstiloColor(wb, "gris", ""), true);
        fila++;

        BigDecimal sumatorioCantidad = new BigDecimal(0);
        BigDecimal sumatorioPorcentaje = new BigDecimal(0);

        for (int z = 0; z < listaUsuarios.size(); z++) {

            String userId = listaUsuarios.get(z).getUserId();
            PersonalData personalData = userService.findById(userId).getPersonalData();

            List<UserParticipation> listaFiltrada = listUP.stream()
                    .filter(participation -> participation.getUserId().equals(userId))
                    .collect(Collectors.toList());

            if (listaFiltrada.size() > 0) {
                BigDecimal cantidadInvertida = listaFiltrada.get(0) != null ? listaFiltrada.get(0).getFunds() : new BigDecimal("0");

                BigDecimal porcentaje = new BigDecimal(0);

                if (companyCapital != null && !companyCapital.equals(new BigDecimal(0))) {

                    porcentaje = cantidadInvertida.multiply(new BigDecimal(100)).divide(companyCapital, 2, RoundingMode.HALF_UP);
                }

                List<String> cabeceraUsuarios = Arrays.asList(personalData.getFirstName() + " " + personalData.getLastName() + " - " + userId, dameNumberFormat().format(cantidadInvertida) + "€", dameNumberFormat().format(porcentaje) + " %");
                imprimirFila(wb, hoja, fila, cabeceraUsuarios, posColumna, dameEstiloGeneral(wb, false, ""), true);
                fila++;

                sumatorioCantidad = sumatorioCantidad.add(cantidadInvertida);
                sumatorioPorcentaje = sumatorioPorcentaje.add(porcentaje);
            }
        }
        fila = fila + 2;

        List<String> sumatorioVoto = Arrays.asList(dameNumberFormat().format(sumatorioCantidad) + "€", dameNumberFormat().format(sumatorioPorcentaje) + " %");
        imprimirFila(wb, hoja, fila, sumatorioVoto, (posColumna + 1), dameEstiloGeneral(wb, true, ""), true);

        List<Double> valoresOcultosVoto = Arrays.asList(sumatorioCantidad.doubleValue(), sumatorioPorcentaje.divide(new BigDecimal(100)).doubleValue());
        imprimirFilaDouble(wb, hoja, posFilaOculta, valoresOcultosVoto, 13, dameEstiloGeneral(wb, false, ""));

        return fila;

    }

    private List<MeetingUserAnswer> dameValorRespuesta(int g, String respuestaVoto, List<MeetingUserAnswer> listUserAnswerFiltrada) {

        List<MeetingUserAnswer> tmp = listUserAnswerFiltrada.stream()
                .filter(userAnswer -> userAnswer.getQuestionAnswerList().get(g).getAnswer().equals(respuestaVoto))
                .collect(Collectors.toList());

        return tmp;
    }

    private XSSFCellStyle dameEstiloColor(XSSFWorkbook wb, String color, String formato) {

        XSSFFont font = wb.createFont();
        font.setBold(true);
        //font.setColor(IndexedColors.BLUE.index); Color de letra
        XSSFCellStyle cellStyle = wb.createCellStyle();

        if (color.equals("gris")) {
            cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(199, 191, 191)));
        } else if (color.equals("verde")) {
            cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(202, 220, 192)));
        } else if (color.equals("negro")) {
            cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(10, 10, 10)));
        }

        if (formato.equals("percetange")) {

            cellStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        } else if (formato.equals("moneda")) {
            cellStyle.setDataFormat(wb.createDataFormat().getFormat("#.###,## €;-#.###,## €"));
        }

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(new java.awt.Color(10, 10, 10)));

        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, new XSSFColor(new java.awt.Color(10, 10, 10)));

        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, new XSSFColor(new java.awt.Color(10, 10, 10)));

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, new XSSFColor(new java.awt.Color(10, 10, 10)));

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private XSSFCellStyle dameEstiloGeneral(XSSFWorkbook wb, Boolean negrita, String formato) {

        XSSFFont font = wb.createFont();
        if (negrita) {
            font.setBold(true);
        }
        //font.setColor(IndexedColors.BLUE.index); Color de letra
        XSSFCellStyle cellStyleGeneral = wb.createCellStyle();

        if (formato.equals("moneda")) {
            cellStyleGeneral.setDataFormat(wb.createDataFormat().getFormat("#.###,## €;-#.###,## €"));
        } else if (formato.equals("percentage")) {
            cellStyleGeneral.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        }

        cellStyleGeneral.setAlignment(HorizontalAlignment.CENTER);
        cellStyleGeneral.setFont(font);
        return cellStyleGeneral;
    }

    @Override
    public void generateDocument(String meetingId)
            throws FileNotFoundException, DocumentException, UrbanitaeBusinessException, InstanceNotFoundException {
        ShareHolderMeeting meeting = this.shareHolderMeetingService.getShareHolderMeeting(meetingId);

        if (meeting.getStatus().equals(MeetingStatus.CLOSED)) {
            MeetingSummary summary = meetingUserAnswerRepository.meetingSummaryResult(meetingId);
            String fileName = meeting.getName() + ".pdf";

            Document document = new Document();
            FileOutputStream fileOutputStream = new FileOutputStream(fileName);
            PdfWriter.getInstance(document, fileOutputStream);

            document.open();
            Font font = FontFactory.getFont(FontFactory.COURIER, 16, BaseColor.BLACK);
            Paragraph chunk = new Paragraph("Resultados de la junta " + meeting.getName() + ".\n", font);
            Paragraph chunkDateStart = new Paragraph("Fecha apertura de la votación " + meeting.getActiveDate() + ".\n",
                    font);
            Paragraph chunkDateEnd = new Paragraph("Fecha cierre de la votación " + meeting.getCloseDate() + ".\n",
                    font);

            document.add(chunk);
            document.add(chunkDateStart);
            document.add(chunkDateEnd);

            summary.getQuestion().stream().forEach(questionSummary -> {
                Paragraph chunkQuestion = new Paragraph("Pregunta " + questionSummary.getQuestion() + ".\n", font);

                try {
                    document.add(chunkQuestion);
                } catch (DocumentException e) {
                    log.error("error writer document {}", ExceptionUtils.getStackTrace(e));
                }

                questionSummary.getResult().stream().forEach(result -> {
                    Paragraph chunkAnswer = new Paragraph(
                            "Respuesta: " + answerKey.getOrDefault(result.getAnswer(), result.getAnswer()) + " - "
                                    + result.getCount() + ".\n",
                            font);
                    try {
                        document.add(chunkAnswer);
                    } catch (DocumentException e) {
                        log.error("error writer document {}", ExceptionUtils.getStackTrace(e));
                    }
                });

            });

            document.close();
            projectUploadFile.uploadFileRestTemplate(meeting, fileName);
        }

    }

    public ShareHolderMeeting saveAdminMeetingUserAnswer(MeetingUserAnswer meetingUserAnswer)
            throws UrbanitaeBusinessException, InstanceNotFoundException {
        ShareHolderMeeting shareHolderMeeting = shareHolderMeetingService
                .getShareHolderMeeting(meetingUserAnswer.getMeetingId());

        shareHolderMeeting.setAdminQuestionAnswerList(meetingUserAnswer.getQuestionAnswerList());

        List<MeetingUserAnswer> userAdminPendingAnswers = this.meetingUserAnswerRepository
                .findByVoteTypeAndMeetingId("delegate_admin", meetingUserAnswer.getMeetingId());

        userAdminPendingAnswers.stream().forEach(delegateAdminAnswer -> {
            delegateAdminAnswer.setQuestionAnswerList(meetingUserAnswer.getQuestionAnswerList());
            this.meetingUserAnswerRepository.save(delegateAdminAnswer);
        });

        checkMeetingClose(meetingUserAnswer, shareHolderMeeting);

        shareHolderMeetingService.save(shareHolderMeeting);
        return shareHolderMeeting;

    }

    private void checkMeetingClose(MeetingUserAnswer meetingUserAnswer, ShareHolderMeeting shareHolderMeeting) {
        long pendingAnswer = shareHolderMeeting.getInvited().stream()
                .filter(investorElement -> investorElement.getCompleteMeeting() == false).count();

        if (pendingAnswer == 0) {
            List<MeetingUserAnswer> answerList = this.meetingUserAnswerRepository
                    .findByMeetingIdIs(meetingUserAnswer.getMeetingId());
            List<MeetingUserAnswer> noOnlineAnswer = answerList.stream().filter(
                    answer -> answer.getQuestionAnswerList() == null || answer.getQuestionAnswerList().isEmpty())
                    .collect(Collectors.toList());

            log.info("pending answers {} ", noOnlineAnswer);

            if (noOnlineAnswer.isEmpty()) {
                shareHolderMeeting.setStatus(MeetingStatus.CLOSED);
                shareHolderMeeting.setCloseDate(new Date());
            }

        }
    }

    @Override
    public MeetingUserAnswer saveDelegateMeetingUserAnswer(MeetingUserAnswer meetingUserAnswer)
            throws UrbanitaeBusinessException, InstanceNotFoundException {
        ShareHolderMeeting shareHolderMeeting = shareHolderMeetingService
                .getShareHolderMeeting(meetingUserAnswer.getMeetingId());

        Investor investor = shareHolderMeeting.getInvited().stream()
                .filter(investorElement -> investorElement.getUserId().compareTo(meetingUserAnswer.getUserId()) == 0)
                .findFirst().orElse(null);
        if (investor != null) {
            if (investor.getCompleteMeeting()) {
                MeetingUserAnswer meetingSave = this.getMeetingUserAnswer(meetingUserAnswer.getMeetingId(),
                        meetingUserAnswer.getUserId());
                checkCanSaveAnswer(meetingUserAnswer, shareHolderMeeting, meetingSave);

                meetingSave.setQuestionAnswerList(meetingUserAnswer.getQuestionAnswerList());
                meetingSave = meetingUserAnswerRepository.save(meetingSave);

                checkMeetingClose(meetingUserAnswer, shareHolderMeeting);

                if (shareHolderMeeting.getStatus().equals(MeetingStatus.CLOSED)) {
                    this.shareHolderMeetingService.save(shareHolderMeeting);
                }

                return meetingSave;

            } else {
                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR,
                        Lists.newArrayList(""), UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR_CODE);
            }
        } else {
            throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR,
                    Lists.newArrayList(""), UrbanitaeExceptionMessages.USER_NOT_IN_MEETING_ERROR_CODE);
        }
    }

    private void checkCanSaveAnswer(MeetingUserAnswer meetingUserAnswer, ShareHolderMeeting shareHolderMeeting,
                                    MeetingUserAnswer meetingSave) throws UrbanitaeBusinessException {
        if (meetingSave.getVoteType().equals("delegate_third")) {
            if (meetingUserAnswer.getQuestionAnswerList() == null
                    || meetingUserAnswer.getQuestionAnswerList().size() != shareHolderMeeting.getQuestions().size()) {
                throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.MEETING_NOT_COMPLETE_ERROR,
                        Lists.newArrayList(""), UrbanitaeExceptionMessages.MEETING_NOT_COMPLETE_ERROR_CODE);
            }
        } else {
            throw new UrbanitaeBusinessException(UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR,
                    Lists.newArrayList(""), UrbanitaeExceptionMessages.INVALID_DATA_MEETING_ERROR_CODE);
        }
    }

}

