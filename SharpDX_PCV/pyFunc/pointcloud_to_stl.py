#!/usr/bin/env python3
"""
点云到STL转换工具
使用Open3D库进行高质量的三角剖分和STL生成
"""

import sys
import os
import json
import numpy as np
import open3d as o3d
from pathlib import Path
import argparse
import time
from scipy.spatial import cKDTree


class PointCloudToSTLConverter:
    """点云到STL转换器"""
    
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        
    def log_progress(self, progress, message):
        """记录进度"""
        if self.progress_callback:
            self.progress_callback(progress, message)
        else:
            print(f"[{progress:3.0f}%] {message}")
    
    def load_txt_files(self, file_paths):
        """加载多个TXT文件"""
        self.log_progress(5, "开始加载TXT文件...")
        
        all_points = []
        
        for i, file_path in enumerate(file_paths):
            self.log_progress(5 + (i * 15 / len(file_paths)), 
                            f"加载文件 {i+1}/{len(file_paths)}: {Path(file_path).name}")
            
            try:
                # 尝试不同的分隔符
                for delimiter in [' ', '\t', ',']:
                    try:
                        data = np.loadtxt(file_path, delimiter=delimiter, comments=['#', '//'])
                        if data.size > 0:
                            break
                    except:
                        continue
                else:
                    # 如果所有分隔符都失败，尝试混合分隔符
                    points = []
                    with open(file_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if not line or line.startswith('#') or line.startswith('//'):
                                continue
                            
                            # 分割行，支持多种分隔符
                            parts = line.replace(',', ' ').replace('\t', ' ').split()
                            if len(parts) >= 2:
                                try:
                                    x = float(parts[0])
                                    y = float(parts[1])
                                    z = float(parts[2]) if len(parts) >= 3 else 0.0
                                    points.append([x, y, z])
                                except ValueError:
                                    continue
                    
                    if points:
                        data = np.array(points)
                    else:
                        print(f"警告: 无法解析文件 {file_path}")
                        continue
                
                # 确保数据是3D的
                if data.ndim == 1:
                    data = data.reshape(1, -1)
                
                if data.shape[1] < 3:
                    # 如果只有2列，添加Z=0
                    if data.shape[1] == 2:
                        z_column = np.zeros((data.shape[0], 1))
                        data = np.hstack([data, z_column])
                    else:
                        print(f"警告: 文件 {file_path} 数据格式不正确")
                        continue
                elif data.shape[1] > 3:
                    # 如果超过3列，只取前3列
                    data = data[:, :3]
                
                all_points.append(data)
                
            except Exception as e:
                print(f"错误: 无法加载文件 {file_path}: {e}")
                continue
        
        if not all_points:
            raise ValueError("没有成功加载任何点云数据")
        
        # 合并所有点
        combined_points = np.vstack(all_points)
        self.log_progress(20, f"成功加载 {len(combined_points)} 个点")
        
        return combined_points
    
    def create_point_cloud(self, points):
        """创建Open3D点云对象"""
        self.log_progress(25, "创建点云对象...")
        
        # 创建Open3D点云
        pcd = o3d.geometry.PointCloud()
        pcd.points = o3d.utility.Vector3dVector(points)
        
        return pcd
    
    def _estimate_spacing(self, points: np.ndarray, sample_size: int = 5000) -> float:
        """估计典型点间距（鲁棒版，优先使用SciPy cKDTree）"""
        try:
            if points is None or len(points) < 2:
                return 1e-2
            pts = np.asarray(points, dtype=float)
            # 去除非有限值
            mask = np.isfinite(pts).all(axis=1)
            pts = pts[mask]
            if len(pts) < 2:
                return 1e-2

            m = min(sample_size, len(pts))
            idx = np.random.choice(len(pts), size=m, replace=False)
            sample = pts[idx]

            # 使用 cKDTree 计算最近邻距离（k=2：自身+最近邻）
            tree = cKDTree(pts, leafsize=64)
            try:
                dists, _ = tree.query(sample, k=2, workers=-1)
            except TypeError:
                # 某些 SciPy 版本没有 workers 参数
                dists, _ = tree.query(sample, k=2)

            nn = dists[:, 1]  # 最近邻距离
            nn = nn[np.isfinite(nn) & (nn > 0)]
            if nn.size == 0:
                return 1e-2
            return float(np.median(nn))
        except Exception:
            # 退化：用包围盒估计（保守）
            try:
                bb = np.ptp(points, axis=0)
                approx = (np.cbrt(np.prod(bb) / max(len(points), 1))) if np.all(bb > 0) else max(bb.max() / 100.0, 1e-2)
                return float(max(approx, 1e-3))
            except Exception:
                return 1e-2

    def preprocess_point_cloud(self, pcd):
        """预处理点云（自适应密度）"""
        self.log_progress(30, "预处理点云...")

        # 估计点间距并做可选下采样（防止过密）
        pts = np.asarray(pcd.points)
        spacing = max(self._estimate_spacing(pts), 1e-4)
        voxel = max(spacing * 0.8, 1e-4)
        # 适度控制点数规模（>30万时也做降采样）
        if len(pts) > 300_000:
            self.log_progress(32, f"点数较多，进行体素下采样 voxel={voxel:.5f}")
            pcd = pcd.voxel_down_sample(voxel)
            pts = np.asarray(pcd.points)  # 更新缓存

        # 移除重复点
        pcd = pcd.remove_duplicated_points()
        self.log_progress(35, f"移除重复点后剩余 {len(pcd.points)} 个点")

        # 自适应离群点移除强度
        nb_neighbors = 30 if len(pcd.points) > 200_000 else 20
        std_ratio = 2.0 if len(pcd.points) > 200_000 else 1.5
        pcd, _ = pcd.remove_statistical_outlier(nb_neighbors=nb_neighbors, std_ratio=std_ratio)
        self.log_progress(40, f"移除离群点后剩余 {len(pcd.points)} 个点")

        # 自适应法向量估计半径
        normal_radius = max(spacing * 3.0, 1e-3)
        max_nn = 80 if len(pcd.points) > 500_000 else (60 if len(pcd.points) > 200_000 else 30)
        try:
            pcd.estimate_normals(search_param=o3d.geometry.KDTreeSearchParamHybrid(radius=normal_radius, max_nn=max_nn))
        except RuntimeError:
            # 退化：使用 KNN 模式
            pcd.estimate_normals(search_param=o3d.geometry.KDTreeSearchParamKNN(knn=min(30, len(pcd.points))))
        self.log_progress(45, f"法向量估计完成 (r={normal_radius:.5f}, nn={max_nn})")

        # 保持法向量一致方向（防止失败）
        try:
            pcd.orient_normals_consistent_tangent_plane(k=min(30, len(pcd.points)))
        except Exception:
            pass

        return pcd
    
    def create_mesh(self, pcd):
        """创建三角网格（自适应参数）"""
        self.log_progress(50, "开始三角剖分...")

        pts = np.asarray(pcd.points)
        spacing = self._estimate_spacing(pts)

        # 优先尝试 Poisson（适合光滑连续表面）
        try:
            depth = 9 if len(pts) < 200_000 else 10
            scale = 1.1
            mesh, _ = o3d.geometry.TriangleMesh.create_from_point_cloud_poisson(
                pcd, depth=depth, width=0, scale=scale, linear_fit=False
            )
            # 可选裁剪：以点云包围盒为界限，适度扩展避免过厚外壳
            try:
                aabb = pcd.get_axis_aligned_bounding_box()
                mesh = mesh.crop(aabb.scale(1.2, aabb.get_center()))
            except Exception:
                pass
            self.log_progress(65, f"Poisson重建完成 (depth={depth})，生成 {len(mesh.triangles)} 个三角形")
        except Exception as e:
            mesh = None
            self.log_progress(55, f"Poisson失败: {e}")

        # 如果 Poisson 网格过大或明显失真，尝试 Alpha-Shape
        if mesh is None or len(mesh.triangles) == 0:
            alpha = max(spacing * 2.5, 5e-4)
            try:
                mesh = o3d.geometry.TriangleMesh.create_from_point_cloud_alpha_shape(pcd, alpha=alpha)
                self.log_progress(65, f"Alpha-Shape成功 (alpha={alpha:.5f})，三角形 {len(mesh.triangles)}")
            except Exception as e2:
                mesh = None
                self.log_progress(60, f"Alpha-Shape失败: {e2}")

        # 如果 Alpha 也不合适，尝试 Ball Pivoting
        if mesh is None or len(mesh.triangles) == 0:
            base = max(spacing, 1e-3)
            radii = [base * r for r in (0.7, 1.0, 1.5, 2.5)]
            try:
                mesh = o3d.geometry.TriangleMesh.create_from_point_cloud_ball_pivoting(
                    pcd, o3d.utility.DoubleVector(radii)
                )
                self.log_progress(65, f"Ball Pivoting成功 (r~{base:.5f})，三角形 {len(mesh.triangles)}")
            except Exception as e3:
                self.log_progress(60, f"Ball Pivoting失败: {e3}")
                raise

        if len(mesh.triangles) == 0:
            raise ValueError("无法生成有效的三角网格")

        return mesh
    
    def postprocess_mesh(self, mesh):
        """后处理网格"""
        self.log_progress(75, "后处理网格...")

        # 移除重复的三角形
        mesh.remove_duplicated_triangles()

        # 移除退化的三角形
        mesh.remove_degenerate_triangles()

        # 移除非流形边
        mesh.remove_non_manifold_edges()

        # 计算顶点法向量（STL保存需要）
        mesh.compute_vertex_normals()

        # 可选：裁剪孤立组件（去除浮动碎片）
        try:
            cc = mesh.cluster_connected_triangles()[0]
            if len(cc) > 0:
                counts = np.bincount(cc)
                keep = counts.argmax()
                mask = (cc == keep)
                triangles = np.asarray(mesh.triangles)
                mesh.triangles = o3d.utility.Vector3iVector(triangles[mask])
                mesh.remove_unreferenced_vertices()
        except Exception:
            pass

        self.log_progress(80, f"后处理完成，最终 {len(mesh.triangles)} 个三角形")

        return mesh
    
    def save_stl(self, mesh, output_path):
        """保存STL文件（确保法向量计算）"""
        self.log_progress(90, f"保存STL文件: {output_path}")

        # 确保输出目录存在
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 确保法向量存在（Open3D保存STL前必需）
        try:
            if not mesh.has_vertex_normals():
                mesh.compute_vertex_normals()
        except Exception:
            mesh.compute_vertex_normals()

        # 作为额外保险，计算三角形法向量
        if not mesh.has_triangle_normals():
            try:
                mesh.compute_triangle_normals()
            except Exception:
                pass

        # 保存为STL格式
        success = o3d.io.write_triangle_mesh(str(output_path), mesh)

        if not success:
            raise ValueError(f"无法保存STL文件到 {output_path}")

        self.log_progress(100, f"STL文件保存成功: {output_path}")

        return str(output_path)
    
    def convert(self, input_files, output_dir, output_name=None):
        """执行完整的转换流程"""
        try:
            # 1. 加载点云数据
            points = self.load_txt_files(input_files)

            # 2. 创建点云对象
            pcd = self.create_point_cloud(points)

            # 3. 预处理
            pcd = self.preprocess_point_cloud(pcd)

            # 4. 创建网格
            mesh = self.create_mesh(pcd)

            # 5. 后处理
            mesh = self.postprocess_mesh(mesh)

            # 6. 构建输出路径
            out_dir = Path(output_dir) if output_dir else Path.cwd()
            out_dir.mkdir(parents=True, exist_ok=True)
            if output_name is None:
                ts = time.strftime('%Y%m%d_%H%M%S')
                output_name = f"pointcloud_{ts}.stl"
            out_path = str(out_dir / output_name)

            # 7. 保存STL
            result_path = self.save_stl(mesh, out_path)
            
            return {
                'success': True,
                'output_path': result_path,
                'original_points': len(points),
                'final_triangles': len(mesh.triangles),
                'final_vertices': len(mesh.vertices)
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='点云到STL转换工具')
    parser.add_argument('input_files', nargs='+', help='输入的TXT文件路径')
    parser.add_argument('-o', '--output-dir', required=False, help='输出目录路径')
    parser.add_argument('--output-name', required=False, help='输出文件名（可选，默认自动命名）')
    parser.add_argument('--json-output', help='JSON格式的结果输出文件')

    args = parser.parse_args()

    # 创建转换器
    converter = PointCloudToSTLConverter()

    # 执行转换
    result = converter.convert(args.input_files, args.output_dir, args.output_name)
    
    # 输出结果
    if args.json_output:
        with open(args.json_output, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
    
    if result['success']:
        print(f"转换成功!")
        print(f"输出文件: {result['output_path']}")
        print(f"原始点数: {result['original_points']:,}")
        print(f"最终三角形数: {result['final_triangles']:,}")
        print(f"最终顶点数: {result['final_vertices']:,}")
        sys.exit(0)
    else:
        print(f"转换失败: {result['error']}")
        sys.exit(1)


if __name__ == '__main__':
    main()
