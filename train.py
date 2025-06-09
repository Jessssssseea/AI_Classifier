import os
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report
import joblib


def clean_text(text):
    return re.sub(r'\s+', ' ', text).strip()


# 设置路径
data_dir = "extracted_texts"  # 假设你已按科目整理好文件夹
categories = [d for d in os.listdir(data_dir) if os.path.isdir(os.path.join(data_dir, d))]

texts = []
labels = []

print("正在加载训练数据...")
for category in categories:
    folder_path = os.path.join(data_dir, category)
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = clean_text(f.read())

            # 使用文件名 + 文件内容作为特征
            filename_base = os.path.splitext(filename)[0]
            combined_text = f"{filename_base} {content}"
            texts.append(combined_text)
            labels.append(category)

        except Exception as e:
            print(f"跳过文件 {file_path}: {e}")

# 向量化
vectorizer = TfidfVectorizer(max_features=10000, ngram_range=(1, 2))
X = vectorizer.fit_transform(texts)

# 划分训练/测试集
X_train, X_test, y_train, y_test = train_test_split(X, labels, test_size=0.2, random_state=42)

# 训练模型
clf = LogisticRegression(max_iter=1000)
clf.fit(X_train, y_train)

# 评估
y_pred = clf.predict(X_test)
print("\n分类报告：")
print(classification_report(y_test, y_pred))

# 保存模型
joblib.dump(clf, "subject_classifier.pkl")
joblib.dump(vectorizer, "tfidf_vectorizer.pkl")

print("\n✅ 模型训练完成，已保存为 subject_classifier.pkl 和 tfidf_vectorizer.pkl")