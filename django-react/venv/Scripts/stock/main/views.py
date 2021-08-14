from django.shortcuts import render
from .models import Post
from rest_framework import generics
from .serializers import PostSerializer
# Create your views here.
def index(request):
    return render(request,'main/index.html')

def blog(request):
    postlist = Post.objects.all()
    return render(request,'main/blog.html',{'postlist':postlist})

def posting(request,pk):
    post = Post.objects.get(pk=pk)
    return render(request,'main/posting/html',{'post':post})

class ListPost(generics.ListCreateAPIView):
    queryset = Post.objects.all()
    serializer_class = PostSerializer
class DetailPost(generics.RetrieveUpdateDestroyAPIView):
    queryset = Post.objects.all()
    serializer_class = PostSerializer