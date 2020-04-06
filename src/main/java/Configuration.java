import java.util.List;

public class Configuration {
    String inputFile;
    String outputFile;
    String outputSheet;
    int numberOfFileToGenerate;
    List<KeepQuesion> keepQuesionList;

    public int getNumberOfFileToGenerate() {
        return numberOfFileToGenerate;
    }

    public void setNumberOfFileToGenerate(int numberOfFileToGenerate) {
        this.numberOfFileToGenerate = numberOfFileToGenerate;
    }

    public String getInputFile() {
        return inputFile;
    }

    public void setInputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    public String getOutputFile() {
        return outputFile;
    }

    public void setOutputFile(String outputFile) {
        this.outputFile = outputFile;
    }

    public String getOutputSheet() {
        return outputSheet;
    }

    public void setOutputSheet(String outputSheet) {
        this.outputSheet = outputSheet;
    }

    public List<KeepQuesion> getKeepQuesionList() {
        return keepQuesionList;
    }

    public void setKeepQuesionList(List<KeepQuesion> keepQuesionList) {
        this.keepQuesionList = keepQuesionList;
    }
}

class KeepQuesion{
    String sheetName;
    long numberOfQuestion;

    public KeepQuesion(String sheetName, long numberOfQuestion) {
        this.sheetName = sheetName;
        this.numberOfQuestion = numberOfQuestion;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public long getNumberOfQuestion() {
        return numberOfQuestion;
    }

    public void setNumberOfQuestion(long numberOfQuestion) {
        this.numberOfQuestion = numberOfQuestion;
    }
}
