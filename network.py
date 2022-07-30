import networkx as nx
import pandas as pd
from pyvis.network import Network


def create_net_lists(SRC_FILE, MEMBER_FILE):
    # ソースファイルよりデータ取得
    df_src = pd.read_excel(SRC_FILE, header=0)
    df_members = pd.read_excel(MEMBER_FILE, header=0)

    # 投稿者= source / @XXX(メンション)対象= target
    source_member_names = df_members['profile_name']
    target_member_names = df_members['mention_name']

    # ソースファイルマージし一部のみ取り出す
    df_merge = pd.merge(df_src, df_members, left_on='投稿者', right_on='profile_name')
    df = df_merge[['投稿者', 'profile_name', 'mention_name', '投稿メッセージ']]

    # 格納用空リスト用意
    data = []

    # 投稿者= sourceを１人ずつ取り出し、
    for i, _source_member_name in enumerate(source_member_names):
        try:
            df_source = df[df['投稿者'] == _source_member_name]
            # 投稿者の名前を@XXXに変換
            source_member_name = df_source.iloc[0, 2]
            print(f'{i} {source_member_name}さんのSlack投稿を抽出...')

            # @XXX(メンション)対象= target
            for target_member_name in target_member_names:
                # 投稿者= sourceのメッセージの中に@XXX(メンション)対象= targetの人物が含まれる行を抽出
                print(f'{i} {source_member_name}さんの -> {target_member_name}さんに対する投稿を抽出...')
                try:
                    relation = df_source[df_source['投稿メッセージ'].str.contains(target_member_name) == True]
                    relation_count = relation.count().sum()

                    datum = {
                        'source': source_member_name,
                        'target': target_member_name,
                        'weight': relation_count
                    }
                    data.append(datum)

                except Exception as e:
                    print(f'{e}')
                    pass
            print(f'{source_member_name}さんのデータ抽出完了！')

        except Exception as e:
            print(f'{e}')
            print(f'{i} 投稿データはありません')

    df = pd.DataFrame(data)
    df.to_excel('src/source_target_weight.xlsx', index=False)

    # Weight = 0(直接的なメッセージ交流なし）除外
    df_net = df[df['weight'] != 0]

    # 名前表示の＠除去
    df_net['source'] = df_net['source'].str.replace('@', '')
    df_net['target'] = df_net['target'].str.replace('@', '')

    return df_net


def create_network(df_net, output_file):
    # エッジリスト
    edges = df_net[['source', 'target', 'weight']].apply(tuple, axis=1).values

    # networkxの形式のデータを生成
    G = nx.Graph()
    G.add_weighted_edges_from(edges)

    # ブラウザ上でインタラクティブに動くネットワーク図を作る
    net = Network('750px', '1200')
    net.from_nx(G)
    net.toggle_physics(True)

    # Weightに応じてエッジの太さを変更
    for i, edge in enumerate(net.edges):
        if abs(edge['weight']) > 10:
            net.edges[i]['width'] = abs(edge['weight']) * 0.01
        else:
            net.edges[i]['width'] = 0.0

    # 設定検討用（コメントアウト外すとGUI設定表示有効化）
    # net.show_buttons(True)  # 全機能
    # net.show_buttons(filter_=['physics','interaction', 'edges', 'nodes'])

    # 調整後の設定（調整したい場合はコメントアウトし、上の行のnet.show_buttons()のコメントアウト外す）
    net.set_options('''
    const options = {
      "nodes": {
        "borderWidth": null,
        "borderWidthSelected": null,
        "opacity": null,
        "font": {
          "size": 14
        },
        "shadow": {
          "enabled": true
        }
      },
      "edges": {
        "color": {
          "inherit": true
        },
        "scaling": {
          "label": {
            "min": null,
            "max": null,
            "maxVisible": null,
            "drawThreshold": null
          }
        },
        "selfReferenceSize": 10,
        "selfReference": {
          "size": 10,
          "angle": 0.7853981633974483
        },
        "smooth": false
      },
      "interaction": {
        "hover": false,
        "multiselect": true,
        "navigationButtons": true
      },
      "physics": {
        "hierarchicalRepulsion": {
          "centralGravity": 0,
          "avoidOverlap": null
        },
        "minVelocity": 0.75,
        "solver": "hierarchicalRepulsion"
      }
    }
    ''')

    # ブラウザにHTML出力
    net.show(output_file)


def main():
    SRC_FILE = 'src/集計.xlsx'
    MEMBER_FILE = 'src/member_profile.xlsx'
    output_file = 'c4b_member_net.html'

    df_net = create_net_lists(SRC_FILE, MEMBER_FILE)
    create_network(df_net, output_file)


if __name__ == '__main__':
    main()