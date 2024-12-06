from dataclasses import dataclass
from typing import List, Optional, Dict
from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException
from datetime import datetime, timedelta

@dataclass
class MonthlySearchVolume:
    """月度搜索量数据类"""
    year_month: str
    monthly_searches: int

@dataclass
class KeywordIdea:
    """关键词创意数据类"""
    text: str
    avg_monthly_searches: int
    competition: str
    competition_index: float
    low_cpc: float  # 首页最低出价
    high_cpc: float  # 首页最高出价
    monthly_searches: List[MonthlySearchVolume]  # 过去12个月的搜索量
    growth_percentage: float  # 年增长百分比
    recent_growth_percentage: float  # 近三个月增长百分比

class KeywordIdeasService:
    """Google Ads关键词创意服务"""
    
    def __init__(self, config_dict: Dict):
        """
        初始化服务
        
        Args:
            config_dict: Google Ads API配置字典，包含必要的认证信息
        """
        self.client = None
        self.customer_id = None
        self.initialize_client(config_dict)
    
    def initialize_client(self, config_dict: Dict) -> None:
        """
        初始化Google Ads客户端
        
        Args:
            config_dict: 配置字典，必须包含以下键：
                        - client_id
                        - client_secret
                        - developer_token
                        - refresh_token
                        - login_customer_id
            
        Raises:
            ValueError: 配置信息不完整
            Exception: 初始化失败
        """
        required_keys = [
            'client_id', 
            'client_secret', 
            'developer_token', 
            'refresh_token',
            'login_customer_id'
        ]
        
        # 验证配置完整性
        missing_keys = [key for key in required_keys if key not in config_dict]
        if missing_keys:
            raise ValueError(f"配置信息不完整，缺少以下字段: {', '.join(missing_keys)}")
            
        try:
            # 添加use_proto_plus配置
            config = dict(config_dict)
            config['use_proto_plus'] = True
            
            self.client = GoogleAdsClient.load_from_dict(config)
            self.customer_id = config['login_customer_id']
                
            if not self.customer_id:
                raise ValueError("配置中未找到login_customer_id")
                
        except Exception as e:
            raise Exception(f"初始化Google Ads客户端失败: {str(e)}")
    
    def get_historical_metrics_batch(self, keywords: List[str], language_id: str = "1000") -> Dict[str, Dict]:
        """
        批量获取关键词的历史指标数据
        
        Args:
            keywords: 关键词列表
            language_id: 语言ID，默认为1000（英语）
            
        Returns:
            Dict[str, Dict]: 关键词到历史指标的映射，包含月度搜索量和其他指标
        """
        if not keywords:
            return {}
            
        try:
            keyword_plan_idea_service = self.client.get_service("KeywordPlanIdeaService")
            googleads_service = self.client.get_service("GoogleAdsService")
            
            request = self.client.get_type("GenerateKeywordHistoricalMetricsRequest")
            request.customer_id = self.customer_id
            request.keywords.extend(keywords)
            request.language = googleads_service.language_constant_path(language_id)
            request.keyword_plan_network = self.client.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
            
            response = keyword_plan_idea_service.generate_keyword_historical_metrics(request=request)
            
            # 创建关键词到指标的映射
            metrics_map = {}
            for result in response.results:
                metrics = result.keyword_metrics
                metrics_map[result.text] = {
                    'keyword': result.text,
                    'monthly_searches': [
                        MonthlySearchVolume(
                            year_month = f"{point.year}-{point.month:02d}",
                            monthly_searches=point.monthly_searches
                        ) for point in metrics.monthly_search_volumes
                    ],
                    'avg_monthly_searches': metrics.avg_monthly_searches,
                    'competition': metrics.competition.name,
                    'competition_index': metrics.competition_index,
                    'low_cpc': metrics.low_top_of_page_bid_micros / 1_000_000,
                    'high_cpc': metrics.high_top_of_page_bid_micros / 1_000_000
                }
                
            return metrics_map
            
        except GoogleAdsException as ex:
            print(
                f'Request with ID "{ex.request_id}" failed with status '
                f'"{ex.error.code().name}" and includes the following errors:'
            )
            for error in ex.failure.errors:
                print(f'\tError with message "{error.message}".')
                if error.location:
                    for field_path_element in error.location.field_path_elements:
                        print(f"\t\tOn field: {field_path_element.field_name}")
            return {}
    
    def calculate_growth_percentage(self, monthly_searches: List[MonthlySearchVolume]) -> float:
        """
        计算年增长百分比
        
        Args:
            monthly_searches: 月度搜索量列表
            
        Returns:
            float: 增长百分比
        """
        if not monthly_searches or len(monthly_searches) < 2:
            return 0.0
            
        # 按年月排序
        sorted_searches = sorted(monthly_searches, key=lambda x: x.year_month)
        first_month = sorted_searches[0].monthly_searches
        last_month = sorted_searches[-1].monthly_searches
        
        if first_month == 0:
            return float('inf') if last_month > 0 else 0.0
            
        return ((last_month - first_month) / first_month) * 100

    def calculate_recent_growth_percentage(self, monthly_searches: List[MonthlySearchVolume]) -> float:
        """
        计算近三个月增长百分比
        
        Args:
            monthly_searches: 月度搜索量列表
            
        Returns:
            float: 近三个月增长百分比
        """
        if not monthly_searches or len(monthly_searches) < 3:
            return 0.0
            
        # 按年月倒序排列，获取最近三个月的数据
        sorted_searches = sorted(monthly_searches, key=lambda x: x.year_month, reverse=True)
        if len(sorted_searches) < 3:
            return 0.0
            
        latest_month = sorted_searches[0].monthly_searches
        third_month = sorted_searches[2].monthly_searches
        
        if third_month == 0:
            return float('inf') if latest_month > 0 else 0.0
            
        return ((latest_month - third_month) / third_month) * 100

    def generate_keyword_ideas(self, keywords: List[str] = None, url: str = None, language_id: str = "1000") -> List[KeywordIdea]:
        """
        获取关键词创意
        
        Args:
            keywords: 关键词列表，可选
            url: 网页URL，可选
            language_id: 语言ID，默认为1000（英语）
            
        Returns:
            List[KeywordIdea]: 关键词创意列表
            
        Raises:
            ValueError: 参数错误
            GoogleAdsException: API调用错误
        """
        if not keywords and not url:
            raise ValueError("关键词列表和URL不能同时为空")
            
        if not self.client or not self.customer_id:
            raise Exception("客户端未初始化")
            
        try:
            keyword_plan_idea_service = self.client.get_service("KeywordPlanIdeaService")
            language_rn = self.client.get_service("GoogleAdsService").language_constant_path(language_id)
            
            request = self.client.get_type("GenerateKeywordIdeasRequest")
            request.customer_id = self.customer_id
            request.language = language_rn
            request.include_adult_keywords = False
            request.keyword_plan_network = self.client.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
            
            # 处理关键词和URL
            keyword_texts = keywords if keywords else []
            
            # 只有URL，没有关键词
            if not keyword_texts and url:
                request.url_seed.url = url
            
            # 只有关键词，没有URL
            elif keyword_texts and not url:
                request.keyword_seed.keywords.extend(keyword_texts)
            
            # 同时有关键词和URL
            elif keyword_texts and url:
                request.keyword_and_url_seed.url = url
                request.keyword_and_url_seed.keywords.extend(keyword_texts)
            
            # 获取关键词创意
            keyword_ideas = keyword_plan_idea_service.generate_keyword_ideas(request=request)
            
            # 提取所有生成的关键词
            generated_keywords = [idea.text for idea in keyword_ideas]
            
            # 检查用户输入的关键词是否在生成的关键词列表中，如果不在则添加
            if keywords:
                for keyword in keywords:
                    if keyword not in generated_keywords:
                        generated_keywords.append(keyword)

            if not generated_keywords:
                raise ValueError("生成的关键词列表为空")
            
            # 批量获取历史数据
            historical_metrics = self.get_historical_metrics_batch(generated_keywords, language_id)
            
            results = []
            for metrics in historical_metrics.values():
                # 从历史数据映射中获取数据
                monthly_searches = metrics.get('monthly_searches', [])
                
                # 计算年增长率和近三个月增长率
                growth_percentage = self.calculate_growth_percentage(monthly_searches)
                recent_growth_percentage = self.calculate_recent_growth_percentage(monthly_searches)
                
                keyword_idea = KeywordIdea(
                    text=metrics['keyword'],
                    avg_monthly_searches=metrics.get('avg_monthly_searches', 0),
                    competition=metrics.get('competition', 'N/A'),
                    competition_index=metrics.get('competition_index', 0),
                    low_cpc=metrics.get('low_cpc', 0),
                    high_cpc=metrics.get('high_cpc', 0),
                    monthly_searches=monthly_searches,
                    growth_percentage=growth_percentage,
                    recent_growth_percentage=recent_growth_percentage
                )
                results.append(keyword_idea)
                
            return results
            
        except GoogleAdsException as ex:
            raise GoogleAdsException(ex.error)
            
        except Exception as e:
            raise Exception(f"获取关键词创意失败: {str(e)}")

# 使用示例
if __name__ == "__main__":
    try:
        # 示例配置
        config = {
            'client_id': 'YOUR_CLIENT_ID',
            'client_secret': 'YOUR_CLIENT_SECRET',
            'developer_token': 'YOUR_DEVELOPER_TOKEN',
            'login_customer_id': 'YOUR_LOGIN_CUSTOMER_ID',
            'refresh_token': 'YOUR_REFRESH_TOKEN',
        }
        
        service = KeywordIdeasService(config)
        keywords = ["python programming"]
        results = service.generate_keyword_ideas(keywords)
        
        for idea in results:
            print(f"关键词: {idea.text}")
            print(f"月平均搜索量: {idea.avg_monthly_searches:,}")
            print(f"竞争度: {idea.competition}")
            print(f"竞争指数: {idea.competition_index}")
            print(f"最低CPC: ${idea.low_cpc:.2f}")
            print(f"最高CPC: ${idea.high_cpc:.2f}")
            print(f"年增长百分比: {idea.growth_percentage:.1f}%")
            print(f"近三个月增长百分比: {idea.recent_growth_percentage:.1f}%")
            print("\n月度搜索量:")
            for monthly in idea.monthly_searches:
                print(f"{monthly.year_month}: {monthly.monthly_searches:,}")
            print("-" * 50)
            
    except Exception as e:
        print(f"错误: {str(e)}")
