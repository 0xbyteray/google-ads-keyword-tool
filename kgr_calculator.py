import requests
from bs4 import BeautifulSoup
import random
import re
from tkinter import messagebox

class KGRCalculator:
    def __init__(self):
        # User-Agent池
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
        ]

    def get_allintitle_count(self, keyword):
        """获取allintitle搜索结果数量"""
        try:
            # 构建搜索URL
            query = f'allintitle:{keyword}'
            url = f'https://www.google.com/search?q={requests.utils.quote(query)}'
            
            # 随机选择User-Agent
            headers = {
                'User-Agent': random.choice(self.user_agents),
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
            }
            
            # 发送请求
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            # 解析结果
            soup = BeautifulSoup(response.text, 'html.parser')
            result_stats = soup.find('div', {'id': 'result-stats'})
            
            if result_stats:
                # 提取数字
                text = result_stats.text
                print("google allintitle 统计:", text)
                
                # 提取"About X results"中的数字
                match = re.search(r'About ([\d,]+) results', text)
                if match:
                    number_str = match.group(1).replace(',', '')  # 移除逗号
                    return int(number_str)
            
            return 0
            
        except Exception as e:
            messagebox.showerror("错误", f"获取allintitle数量时出错: {str(e)}")
            return 0

    def calculate(self, keyword, monthly_searches, avg_monthly_searches):
        """计算KGR值
        
        Args:
            keyword: 关键词
            monthly_searches: 最近一个月的搜索量
            avg_monthly_searches: 月平均搜索量
            
        Returns:
            tuple: (kgr_avg, kgr_latest, allintitle_count) KGR平均值、最新KGR值和allintitle数量
        """
        # 获取allintitle数量
        allintitle_count = self.get_allintitle_count(keyword)
            
        # 计算基于平均搜索量的KGR
        if avg_monthly_searches == 0:
            kgr_avg = float('inf')
        else:
            kgr_avg = allintitle_count / avg_monthly_searches
            
        # 计算基于最新月搜索量的KGR
        if monthly_searches == 0:
            kgr_latest = float('inf')
        else:
            kgr_latest = allintitle_count / monthly_searches
            
        return kgr_avg, kgr_latest, allintitle_count
        
def main():
    """测试用例"""
    calculator = KGRCalculator()
    
    keyword = "python programming"
    monthly_searches = 1000
    avg_monthly_searches = 500
    
    kgr_avg, kgr_latest, allintitle_count = calculator.calculate(keyword, monthly_searches, avg_monthly_searches)
    
    print(f"{keyword}的平均KGR值为: {kgr_avg:.2f}, 最新KGR值为: {kgr_latest:.2f}, allintitle数量为: {allintitle_count}")
    

if __name__ == "__main__":
    main()
