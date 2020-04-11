import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

public class App {
    public static void main(String[] args) throws Exception {
        try {
            generateFile();
        }catch (Exception ex){
            ex.printStackTrace();
            System.out.println(ex.toString());
        }
    }

    private static List<Question> removeItem(List<Question> questionList, int numberOfKeep) {
        int numberOfRemoval = questionList.size() - numberOfKeep;
        System.out.println(numberOfRemoval);
        Random random = new Random();
        for (int i = 0; i < numberOfRemoval; i++) {
            System.out.println(questionList.size());
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
                ExcelReader.writeExce3(masterList, config.getOutputSheet(), outputFileName);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            System.out.println("Error: \n" + ex.toString());
        }
    }
}

