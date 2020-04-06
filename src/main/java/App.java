import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

public class App {
    public static void main(String[] args) throws Exception {

        generateFile();


//        List<Data> master = ExcelReader.readExcel();
//
//        Configuration config = JSONFileReader.readConfiguration();
//        final List<Question> masterList = new ArrayList<>();
//
//        for (Data listQuestion : master) {
//            String sheetName = listQuestion.getSheetName();
//            int numberOfKeep = 0;
//            for (KeepQuesion keep : config.getKeepQuesionList()) {
//                if (keep.getSheetName().equals(sheetName)) {
//                    numberOfKeep = (int) keep.getNumberOfQuestion();
//                    break;
//                }
//            }
//            removeItem(listQuestion.getQuestionList(), numberOfKeep);
//            listQuestion.getQuestionList().forEach(e -> masterList.add(e));
//        }
//
//        Collections.shuffle(masterList);
//
//        ExcelReader.writeExcel2(masterList, config.getOutputSheet(), config.getOutputFile());
    }

    private static List<Question> removeItem(List<Question> questionList, int numberOfKeep) {
        int numberOfRemoval = questionList.size() - numberOfKeep;
        Random random = new Random();
        for (int i = 0; i < numberOfRemoval; i++) {
            int removeIndex = random.nextInt(questionList.size() - 1);
            questionList.remove(removeIndex);
        }
        return questionList;
    }

    private static void generateFile() {
        try {
            Configuration config = JSONFileReader.readConfiguration();
            for (int i = 0; i < config.getNumberOfFileToGenerate(); i++) {
                List<Data> masterData = ExcelReader.readExcel();
                List<Question> masterList = new ArrayList<>();

                for (Data data : masterData) {
                    String sheetName = data.getSheetName();
                    int numberOfKeep = 0;
                    for (KeepQuesion keep : config.getKeepQuesionList()) {
                        if (keep.getSheetName().equals(sheetName)) {
                            numberOfKeep = (int) keep.getNumberOfQuestion();
                            break;
                        }
                    }
                    removeItem(data.getQuestionList(), numberOfKeep);
                    data.getQuestionList().forEach(e -> masterList.add(e));
                }
                Collections.shuffle(masterList);
                String outputFileName = String.format(config.getOutputFile(), i);
                ExcelReader.writeExcel2(masterList, config.getOutputSheet(), outputFileName);
            }
        } catch (Exception ex) {
            System.out.println("Error: \n" + ex.toString());
        }
    }
}

