import java.util.ArrayList;
import java.util.List;

public class Data {
    String sheetName;
    List<Question> questionList = new ArrayList();

    public Data(String sheetName, List<Question> questionList) {
        this.sheetName = sheetName;
        this.questionList = questionList;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void addQuestion(Question question){
        questionList.add(question);
    }

    public List<Question> getQuestionList() {
        return questionList;
    }
}
